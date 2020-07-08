VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_GesFia_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   2505
   ClientTop       =   1755
   ClientWidth     =   14130
   Icon            =   "OpeTra_frm_037.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14115
      _Version        =   65536
      _ExtentX        =   24897
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
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
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
         Begin VB.CommandButton cmd_Evalua 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_037.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Cargar Operación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13380
            Picture         =   "OpeTra_frm_037.frx":08D6
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
            Width           =   6345
            _Version        =   65536
            _ExtentX        =   11192
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Cartas Fianza"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            Picture         =   "OpeTra_frm_037.frx":0D18
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7245
         Left            =   30
         TabIndex        =   6
         Top             =   1440
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
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   60
            TabIndex        =   7
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
         Begin Threed.SSPanel pnl_Tit_FecEmi 
            Height          =   285
            Left            =   8430
            TabIndex        =   10
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Emisión"
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
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   30
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_FecVct 
            Height          =   285
            Left            =   9930
            TabIndex        =   12
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Vcto."
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
         Begin Threed.SSPanel pnl_Tit_Situac 
            Height          =   285
            Left            =   11430
            TabIndex        =   13
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
Attribute VB_Name = "frm_GesFia_01"
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
   moddat_g_str_NumOpe = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 2) & Mid(grd_Listad.Text, 8, 5)
   
   grd_Listad.Col = 1
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
         
   grd_Listad.Col = 2
   moddat_g_str_NomCli = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgAct = 1
   
   frm_GesFia_02.Show 1
   
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
   grd_Listad.ColWidth(0) = 1500
   grd_Listad.ColWidth(1) = 1500
   grd_Listad.ColWidth(2) = 5355
   grd_Listad.ColWidth(3) = 1500
   grd_Listad.ColWidth(4) = 1520
   grd_Listad.ColWidth(5) = 2205
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar()
   Dim r_str_NumDoc     As String
   Dim r_int_TipDoc     As Integer
   
   g_str_Parame = "SELECT * FROM CRE_HIPFIA A, CRE_HIPMAE B WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = HIPFIA_NUMOPE AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2  AND "
   g_str_Parame = g_str_Parame & "HIPFIA_SITUAC <> 2 AND "
   g_str_Parame = g_str_Parame & "HIPFIA_SITUAC <> 3 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPFIA_EMIFIA ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
         g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & g_rst_Princi!HIPFIA_NUMOPE & "' "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            r_int_TipDoc = g_rst_Genera!HIPMAE_TDOCLI
            r_str_NumDoc = Trim(g_rst_Genera!HIPMAE_NDOCLI & "")
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
      
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = gf_Formato_NumOpe(g_rst_Princi!HIPFIA_NUMOPE)
            
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(r_int_TipDoc) & "-" & Trim(r_str_NumDoc)
            
         grd_Listad.Col = 2
         grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(r_int_TipDoc), Trim(r_str_NumDoc))
            
         grd_Listad.Col = 3
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPFIA_EMIFIA))
            
         grd_Listad.Col = 4
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPFIA_VCTFIA))
            
         grd_Listad.Col = 5
         grd_Listad.Text = moddat_gf_Consulta_ParDes("032", CStr(g_rst_Princi!HIPFIA_SITUAC))
            
         grd_Listad.Col = 6
         grd_Listad.Text = CStr(g_rst_Princi!HIPFIA_EMIFIA)
            
         grd_Listad.Col = 7
         grd_Listad.Text = CStr(g_rst_Princi!HIPFIA_VCTFIA)
            
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   
   If grd_Listad.Rows = 0 Then
      cmd_Evalua.Enabled = False
      
      MsgBox "No se encontraron Cartas Fianza registradas.", vbInformation, modgen_g_str_NomPlt
   Else
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   
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

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecEmi_Click()
   If Len(Trim(pnl_Tit_FecEmi.Tag)) = 0 Or pnl_Tit_FecEmi.Tag = "D" Then
      pnl_Tit_FecEmi.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "N")
   Else
      pnl_Tit_FecEmi.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "N-")
   End If
End Sub

Private Sub pnl_Tit_FecVct_Click()
   If Len(Trim(pnl_Tit_FecVct.Tag)) = 0 Or pnl_Tit_FecVct.Tag = "D" Then
      pnl_Tit_FecVct.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 7, "N")
   Else
      pnl_Tit_FecVct.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 7, "N-")
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

Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_Situac_Click()
   If Len(Trim(pnl_Tit_Situac.Tag)) = 0 Or pnl_Tit_Situac.Tag = "D" Then
      pnl_Tit_Situac.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_Situac.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub
