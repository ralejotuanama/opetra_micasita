VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Ges_CreHip_12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   1620
   ClientTop       =   1755
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_147.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
      _ExtentY        =   8705
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
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   1
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1349
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1440
            TabIndex        =   2
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1440
            TabIndex        =   3
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   4
            Top             =   390
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2625
         Left            =   30
         TabIndex        =   6
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   4630
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Banco"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   3090
            TabIndex        =   8
            Top             =   60
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Carta"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   5370
            TabIndex        =   9
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   6870
            TabIndex        =   10
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   8370
            TabIndex        =   11
            Top             =   60
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   9660
            TabIndex        =   12
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
            Height          =   2235
            Left            =   30
            TabIndex        =   13
            Top             =   360
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   3942
            _Version        =   393216
            Rows            =   30
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Left            =   600
            TabIndex        =   19
            Top             =   30
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Gestión de Crédito Hipotecario"
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
            Height          =   315
            Left            =   600
            TabIndex        =   20
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cartas Fianza"
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
            Picture         =   "OpeTra_frm_147.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   15
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Left            =   10890
            Picture         =   "OpeTra_frm_147.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueFia 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_147.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Renovación de Carta Fianza"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_LibFia 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_147.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Liberación de Carta Fianza"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_CreHip_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public l_str_Estado        As String
Private l_int_Situac       As Integer

Private Sub cmd_LibFia_Click()
   Dim r_int_Situac     As Integer

   If grd_Listad.Row > 0 Then
      MsgBox "Este Registro no podra ser modificado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Row = 0
   
   grd_Listad.Col = 7
   moddat_g_str_BanFia = grd_Listad.Text
   
   grd_Listad.Col = 1
   moddat_g_str_NumFia = grd_Listad.Text
   
   grd_Listad.Col = 8
   moddat_g_str_FecFia = grd_Listad.Text
   
   grd_Listad.Col = 6
   r_int_Situac = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If l_int_Situac = 2 Then
      If moddat_g_int_TipGar <> 1 Then
         MsgBox "No se puede liberar la Carta Fianza mientras no se haya registrado la Hipoteca.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   If r_int_Situac = 2 Then
      MsgBox "La Carta Fianza ya fue liberada.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Ges_CreHip_14.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_NueFia_Click()
   Dim r_int_Situac     As Integer

'   If moddat_g_int_TipGar = 1 Or moddat_g_int_TipGar = 2 Then
'      MsgBox "Operación ya tiene registrada la HIPOTECA.", vbExclamation, modgen_g_str_NomPlt
'      Exit Sub
'   End If
   
   If grd_Listad.Row > 0 Then
      MsgBox "Este Registro no podra ser modificado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Row = 0
   
   grd_Listad.Col = 7
   moddat_g_str_BanFia = grd_Listad.Text
   
   grd_Listad.Col = 1
   moddat_g_str_NumFia = grd_Listad.Text
   
   grd_Listad.Col = 8
   moddat_g_str_FecFia = grd_Listad.Text
   
   grd_Listad.Col = 6
   r_int_Situac = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgAct = 1
   
   l_str_Estado = "N"

   frm_Ges_CreHip_13.Show 1
   
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
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Buscar
   Call fs_DatCre
   
'   'Validacion para los creditos extornados y cancelados
'   If moddat_g_int_Situac = 7 Or moddat_g_int_Situac = 9 Then
'      cmd_LibFia.Enabled = False
'      cmd_NueFia.Enabled = False
'   Else
'      cmd_LibFia.Enabled = True
'      cmd_NueFia.Enabled = True
'   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 3045:      grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColWidth(1) = 2295:      grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColWidth(2) = 1515:      grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColWidth(3) = 1485:      grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColWidth(4) = 1305:      grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColWidth(5) = 1485:      grd_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
End Sub

Public Sub fs_Buscar()
   moddat_g_str_FecIni = ""

   g_str_Parame = "SELECT * FROM CRE_HIPFIA WHERE "
   g_str_Parame = g_str_Parame & "HIPFIA_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "ORDER BY HIPFIA_EMIFIA DESC"
   
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
         grd_Listad.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPFIA_BANFIA & "")
            
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(g_rst_Princi!HIPFIA_NUMFIA & "")
            
         grd_Listad.Col = 2
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPFIA_EMIFIA))
         
         If Len(Trim(moddat_g_str_FecIni)) = 0 Then
            moddat_g_str_FecIni = gf_FormatoFecha(CStr(g_rst_Princi!HIPFIA_EMIFIA))
         End If
            
         grd_Listad.Col = 3
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPFIA_VCTFIA))
            
         grd_Listad.Col = 4
         grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!HIPFIA_MONFIA))
         
         grd_Listad.Col = 5
         grd_Listad.Text = Format(g_rst_Princi!HIPFIA_IMPFIA, "###,##0.00")
            
         grd_Listad.Col = 6
         grd_Listad.Text = CStr(g_rst_Princi!HIPFIA_SITUAC)
         
         grd_Listad.Col = 7
         grd_Listad.Text = g_rst_Princi!HIPFIA_BANFIA & ""
         
         grd_Listad.Col = 8
         grd_Listad.Text = CStr(g_rst_Princi!HIPFIA_EMIFIA)
         
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   
   If grd_Listad.Rows > 0 Then
      Call gs_UbiIniGrid(grd_Listad)
   Else
      cmd_NueFia.Enabled = False
      cmd_LibFia.Enabled = False
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   'Obteniendo datos del Maestro de Créditos
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   moddat_g_int_TipGar = g_rst_Princi!HIPMAE_TIPGAR
   l_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Obteniendo datos del Maestro de Hipotecas
   moddat_g_str_FecHip = ""
   
   g_str_Parame = "SELECT * FROM CRE_HIPGAR WHERE "
   g_str_Parame = g_str_Parame & "HIPGAR_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      moddat_g_str_FecHip = gf_FormatoFecha(CStr(g_rst_Princi!HIPGAR_FECCON))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   'If grd_Listad.Row <= 0 Then
      'Valida que no tenga registrada la garantia
      If moddat_g_int_TipGar = 1 Or moddat_g_int_TipGar = 2 Then
         MsgBox "Operación ya tiene registrada la HIPOTECA.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If

      l_str_Estado = "M"
      frm_Ges_CreHip_13.Show 1
   'Else
   '   MsgBox "Este Registro no podra ser modificado.", vbExclamation, modgen_g_str_NomPlt
   'End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub
