VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_MntCli_56 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   4425
   ClientTop       =   4590
   ClientWidth     =   11655
   Icon            =   "OpeTra_frm_165.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   5794
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   1245
         Left            =   30
         TabIndex        =   1
         Top             =   1980
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   2196
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
            Height          =   855
            Left            =   60
            TabIndex        =   13
            Top             =   360
            Width           =   11475
            _ExtentX        =   20241
            _ExtentY        =   1508
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   60
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Orden Actividad Económica"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   3210
            TabIndex        =   15
            Top             =   60
            Width           =   7995
            _Version        =   65536
            _ExtentX        =   14102
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción Actividad Económica"
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
         TabIndex        =   2
         Top             =   30
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   255
            Left            =   660
            TabIndex        =   8
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   255
            Left            =   660
            TabIndex        =   9
            Top             =   330
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Actividades Económicas"
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
            Picture         =   "OpeTra_frm_165.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   3
         Top             =   1470
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   767
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1950
            TabIndex        =   4
            Top             =   60
            Width           =   9615
            _Version        =   65536
            _ExtentX        =   16960
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1 - 07522154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin VB.Label Label1 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   5
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   6
         Top             =   750
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_165.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Borrar Actividad Económica"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_165.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Modificar Actividad Económica"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_165.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Nueva Actividad Económica"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10980
            Picture         =   "OpeTra_frm_165.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Datos del Crédito"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_56"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()
   If grd_Listad.Rows = 2 Then
      MsgBox "Sólo se pueden ingresar dos actividades económicas.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If grd_Listad.Rows = 0 Then
      modmip_g_int_OrdAct = 1
   ElseIf grd_Listad.Rows = 1 Then
      modmip_g_int_OrdAct = 2
   End If
   
   modmip_g_int_FlgGrb_1 = 1
   modmip_g_int_FlgAct_1 = 1
   
   frm_MntCli_66.Show 1
   
   If modmip_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      
      Call fs_Buscar
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Borrar_Click()
   Dim r_int_OrdAct     As Integer
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 2
   r_int_OrdAct = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If r_int_OrdAct = 1 And grd_Listad.Rows = 2 Then
      MsgBox "No puede eliminar la Actividad Económica Principal, debe primero eliminar la Actividad Económica Secundaria.", vbExclamation, modgen_g_str_NomPlt
      
      Exit Sub
   ElseIf r_int_OrdAct = 1 And grd_Listad.Rows = 1 Then
      If MsgBox("¿Está seguro de eliminar la Actividad Económica Principal?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   ElseIf r_int_OrdAct = 2 Then
      If MsgBox("¿Está seguro de eliminar la Actividad Económica Secundaria?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   End If
   
   Screen.MousePointer = 11
   
   'Instrucción SQL
   g_str_Parame = "DELETE FROM CLI_ACTECO WHERE "
   
   If modmip_g_int_TipCli = 1 Then
      g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_NumDoc) & "' AND "
   Else
      g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_CygNDo) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(r_int_OrdAct) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If r_int_OrdAct = 1 And grd_Listad.Rows = 1 Then
      'Actualizar en Maestro de Clientes
      g_str_Parame = "USP_CLI_DATGEN_ACTECOPRI ("
      
      If modmip_g_int_TipCli = 1 Then
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      End If
      
      g_str_Parame = g_str_Parame & "0, "          'Código Actividad Económica Principal
      g_str_Parame = g_str_Parame & "0, "          'Código CIIU
      g_str_Parame = g_str_Parame & "0, "          'Tipo DOI
      g_str_Parame = g_str_Parame & "'', "         'Numero DOI
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_CLI_ACTECO_AGREGA.", vbCritical, modgen_g_str_NomPlt
         
         Exit Sub
      End If
   End If
   
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Editar_Click()
   Dim r_int_CodAct     As Integer
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 2
   modmip_g_int_OrdAct = CInt(grd_Listad.Text)
   
   grd_Listad.Col = 3
   r_int_CodAct = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   modmip_g_int_FlgGrb_1 = 2
   modmip_g_int_FlgAct_1 = 1
   
   If modmip_g_int_PaiRes = 4028 Then
      Select Case r_int_CodAct
         Case 11: frm_MntCli_57.Show 1
         Case 21: frm_MntCli_58.Show 1
         Case 31: frm_MntCli_59.Show 1
         Case 41: frm_MntCli_60.Show 1
         Case 51: frm_MntCli_61.Show 1
         Case 61: frm_MntCli_62.Show 1
      End Select
   Else
      Select Case r_int_CodAct
         Case 11: frm_MntCli_63.Show 1
         Case 21: frm_MntCli_64.Show 1
         Case 31: frm_MntCli_65.Show 1
         Case 61: frm_MntCli_62.Show 1
      End Select
   End If
   
   If modmip_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      
      Call fs_Buscar
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Salida_Click()
   If grd_Listad.Rows = 0 Then
      If MsgBox("No ha ingresado ninguna Actividad Económica. ¿Está seguro de salir?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   End If
   
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   If modmip_g_int_TipCli = 1 Then
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   Else
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli & " (" & CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom & ")"
   End If
   
   Call fs_Inicio
   Call fs_Buscar
   
   'Buscar Actividades Económicas
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   grd_Listad.ColWidth(0) = 3105
   grd_Listad.ColWidth(1) = 7995
   
   grd_Listad.ColWidth(2) = 0
   grd_Listad.ColWidth(3) = 0
End Sub

Private Sub fs_Buscar()
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False

   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   
   If modmip_g_int_TipCli = 1 Then
      g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_NumDoc & "' ORDER BY ACTECO_ORDACT ASC"
   Else
      g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_CygNDo & "' ORDER BY ACTECO_ORDACT ASC"
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct))
      
         grd_Listad.Col = 1
         grd_Listad.Text = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ACTECO_CODACT))
         
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!ActEco_OrdAct)
      
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!ACTECO_CODACT)
      
         g_rst_Princi.MoveNext
      Loop
      
      grd_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad)
   
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

