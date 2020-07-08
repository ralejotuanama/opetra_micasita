VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Desemb_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8745
   ClientLeft      =   4725
   ClientTop       =   2880
   ClientWidth     =   15135
   Icon            =   "OpeTra_frm_033.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      _Version        =   65536
      _ExtentX        =   26696
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
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
            Height          =   315
            Left            =   690
            TabIndex        =   13
            Top             =   30
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel SSPanel68 
            Height          =   315
            Left            =   690
            TabIndex        =   14
            Top             =   330
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Desembolso"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_033.frx":000C
            Stretch         =   -1  'True
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7245
         Left            =   30
         TabIndex        =   2
         Top             =   1440
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   6855
            Left            =   30
            TabIndex        =   8
            Top             =   330
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   30
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   3495
            _Version        =   65536
            _ExtentX        =   6165
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
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
            Left            =   3540
            TabIndex        =   4
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
            Left            =   6540
            TabIndex        =   5
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
            Left            =   8040
            TabIndex        =   6
            Top             =   60
            Width           =   5145
            _Version        =   65536
            _ExtentX        =   9075
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
         Begin Threed.SSPanel pnl_Tit_FecGen 
            Height          =   285
            Left            =   13170
            TabIndex        =   7
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Generación"
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
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   5040
            TabIndex        =   9
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
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
            Left            =   14430
            Picture         =   "OpeTra_frm_033.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Evalua 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_033.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Abrir Operación"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Desemb_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Evalua_Click()

MsgBox ("le hemos dado doble click")

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_NomPrd = grd_Listad.Text

   grd_Listad.Col = 1
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
   
   grd_Listad.Col = 2
   moddat_g_str_NumOpe = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 2) & Mid(grd_Listad.Text, 8, 5)
   
   grd_Listad.Col = 3
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
         
   grd_Listad.Col = 4
   moddat_g_str_NomCli = grd_Listad.Text
   
   grd_Listad.Col = 5
   moddat_g_str_FecIng = grd_Listad.Text
   
   grd_Listad.Col = 6
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 7
   moddat_g_str_CodSub = grd_Listad.Text
   
   grd_Listad.Col = 9
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
'   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
'      MsgBox "Debe registrar el Tipo de Cambio pra la moneda " & moddat_gf_Consulta_ParDes("204", CStr(2)) & ".", vbExclamation, modgen_g_str_NomPlt
'      Exit Sub
'   End If
'   If moddat_g_int_TipMon <> 1 Then
'      If moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon) = 0 Then
'         MsgBox "Debe registrar el Tipo de Cambio pra la moneda " & moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon)) & ".", vbExclamation, modgen_g_str_NomPlt
'         Exit Sub
'      End If
'   End If
   
   moddat_g_int_FlgAct = 1
   
   frm_Desemb_22.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Inicia
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
   grd_Listad.ColWidth(0) = 3485
   grd_Listad.ColWidth(1) = 1505
   grd_Listad.ColWidth(2) = 1505
   grd_Listad.ColWidth(3) = 1505
   grd_Listad.ColWidth(4) = 5135
   grd_Listad.ColWidth(5) = 1505
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_CODPRD, HIPMAE_NUMSOL, HIPMAE_NUMOPE, HIPMAE_TDOCLI, HIPMAE_NDOCLI, "
   g_str_Parame = g_str_Parame & "       HIPMAE_FECACT, HIPMAE_CODPRD, HIPMAE_CODSUB, HIPMAE_MONEDA, PRODUC_DESCRI, "
   g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOM_CLIENTE "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE  "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC ON PRODUC_CODIGO = HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPMAE_TDOCLI AND DATGEN_NUMDOC = HIPMAE_NDOCLI "
   'g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE IN ('0071200079','0071300002','0071300013')  "
   g_str_Parame = g_str_Parame & " WHERE (HIPMAE_SITUAC = 1 OR HIPMAE_SITUAC = 2) "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_TIPGAR <> 1 "
   g_str_Parame = g_str_Parame & " ORDER BY HIPMAE_FECACT DESC"
   
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
         grd_Listad.Text = Trim(g_rst_Princi!PRODUC_DESCRI)
         
         grd_Listad.Col = 1
         grd_Listad.Text = Left(g_rst_Princi!hipmae_numsol, 3) & "-" & Mid(g_rst_Princi!hipmae_numsol, 4, 3) & "-" & Mid(g_rst_Princi!hipmae_numsol, 7, 2) & "-" & Right(g_rst_Princi!hipmae_numsol, 4)
         
         grd_Listad.Col = 2
         grd_Listad.Text = Left(g_rst_Princi!HIPMAE_NUMOPE, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMOPE, 4, 2) & "-" & Right(g_rst_Princi!HIPMAE_NUMOPE, 5)
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
         
         grd_Listad.Col = 4
         grd_Listad.Text = g_rst_Princi!NOM_CLIENTE
         
         grd_Listad.Col = 5
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT))
         
         grd_Listad.Col = 6
         grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_CODPRD & "")
         
         grd_Listad.Col = 7
         grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_CODSUB & "")
         
         grd_Listad.Col = 8
         grd_Listad.Text = g_rst_Princi!HIPMAE_FECACT
         
         grd_Listad.Col = 9
         grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_MONEDA)
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   
   If grd_Listad.Rows = 0 Then
      cmd_Evalua.Enabled = False
      MsgBox "No se encontraron Operaciones Pendientes de Desembolso.", vbInformation, modgen_g_str_NomPlt
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

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecGen_Click()
   If Len(Trim(pnl_Tit_FecGen.Tag)) = 0 Or pnl_Tit_FecGen.Tag = "D" Then
      pnl_Tit_FecGen.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "N")
   Else
      pnl_Tit_FecGen.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Len(Trim(pnl_Tit_NumSol.Tag)) = 0 Or pnl_Tit_NumSol.Tag = "D" Then
      pnl_Tit_NumSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

