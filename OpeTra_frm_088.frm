VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_EmpPer_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7380
   ClientLeft      =   4455
   ClientTop       =   1920
   ClientWidth     =   8190
   Icon            =   "OpeTra_frm_088.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   13044
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
         Height          =   5895
         Left            =   30
         TabIndex        =   1
         Top             =   1440
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
         _ExtentY        =   10398
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
         Begin Threed.SSPanel pnl_Tit_RazSoc 
            Height          =   285
            Left            =   1500
            TabIndex        =   2
            Top             =   60
            Width           =   6315
            _Version        =   65536
            _ExtentX        =   11139
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Razón Social"
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
         Begin Threed.SSPanel pnl_Tit_CodEmp 
            Height          =   285
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Empresa"
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
            Height          =   5505
            Left            =   30
            TabIndex        =   4
            Top             =   360
            Width           =   8025
            _ExtentX        =   14155
            _ExtentY        =   9710
            _Version        =   393216
            Rows            =   12
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   30
         TabIndex        =   5
         Top             =   750
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
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
         Begin VB.CommandButton cmd_DatEmp 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_088.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Datos Empresa"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   3660
            Picture         =   "OpeTra_frm_088.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Imprimir Listado"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_PerCon 
            Height          =   585
            Left            =   3060
            Picture         =   "OpeTra_frm_088.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Personas de Contacto"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Perito 
            Height          =   585
            Left            =   2460
            Picture         =   "OpeTra_frm_088.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Empleados"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_088.frx":132C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_088.frx":1636
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_088.frx":1940
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7500
            Picture         =   "OpeTra_frm_088.frx":1C4A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   480
            Left            =   630
            TabIndex        =   8
            Top             =   90
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Empresas"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   7650
            Top             =   60
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   90
            Picture         =   "OpeTra_frm_088.frx":208C
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_EmpPer_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgAct = 1
   moddat_g_int_FlgGrb = 1
   
   frm_EmpPer_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar(moddat_g_int_TipRec)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar la Empresa?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Obteniendo Información del Registro
   g_str_Parame = "USP_BORRAR_MNT_PARDES_ITEM ("
   g_str_Parame = g_str_Parame & "'507', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
   
   Call fs_Buscar(moddat_g_int_TipRec)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_DatEmp_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
         
   grd_Listad.Col = 1
   moddat_g_str_DesGrp = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
      
   If moddat_g_int_TipRec = 3 Then
      frm_EmpPer_07.Show 1
   ElseIf moddat_g_int_TipRec = 2 Then
      frm_EmpPer_09.Show 1
   ElseIf moddat_g_int_TipRec = 1 Then
      frm_EmpPer_08.Show 1
   End If
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgAct = 1
   moddat_g_int_FlgGrb = 2
   
   frm_EmpPer_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar(moddat_g_int_TipRec)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Imprim_Click()
   If MsgBox("¿Está seguro de Imprimir el reporte?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "MNT_PARDES"
   If moddat_g_int_TipRec = 3 Then
      crp_Imprim.SelectionFormula = "{MNT_PARDES.PARDES_CODGRP} = '509' AND {MNT_PARDES.PARDES_CODITE} <> '000000' "
   Else
      crp_Imprim.SelectionFormula = "{MNT_PARDES.PARDES_CODGRP} = '507' AND {MNT_PARDES.PARDES_CODITE} <> '000000' "
   End If
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "MNT_EMPPER_01.RPT"
   crp_Imprim.Destination = crptToWindow
   
   If moddat_g_int_TipRec = 1 Then
      crp_Imprim.ParameterFields(0) = "p_Titulo;" & "REPORTE DE EMPRESAS DE PERITAJE" & ";True"
   ElseIf moddat_g_int_TipRec = 2 Then
      crp_Imprim.ParameterFields(0) = "p_Titulo;" & "REPORTE DE EMPRESAS DE SEGUROS" & ";True"
   ElseIf moddat_g_int_TipRec = 3 Then
      crp_Imprim.ParameterFields(0) = "p_Titulo;" & "REPORTE DE EMPRESAS DE NOTARIA" & ";True"
   End If
   
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_PerCon_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
         
   grd_Listad.Col = 1
   moddat_g_str_DesGrp = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_EmpPer_05.Show 1
End Sub

Private Sub cmd_Perito_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
         
   grd_Listad.Col = 1
   moddat_g_str_DesGrp = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   frm_EmpPer_03.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar(moddat_g_int_TipRec)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1455:      grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColWidth(1) = 6315:      grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   If moddat_g_int_TipRec = 1 Then
      pnl_TitPri.Caption = "Empresas de Peritaje"
   ElseIf moddat_g_int_TipRec = 2 Then
      pnl_TitPri.Caption = "Empresas de Seguros"
   ElseIf moddat_g_int_TipRec = 3 Then
      pnl_TitPri.Caption = "Empresas de Notarias"
   End If
End Sub

Private Sub fs_Buscar(p_TipEmp As Integer)
Dim r_str_Parame   As String
Dim r_rst_Genera   As ADODB.Recordset

   cmd_Agrega.Enabled = False
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_Perito.Enabled = False
   cmd_PerCon.Enabled = False
   cmd_Imprim.Enabled = False
   
   grd_Listad.Enabled = False
   Call gs_LimpiaGrid(grd_Listad)
   
   If p_TipEmp = 1 Or p_TipEmp = 3 Then
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT *  "
      r_str_Parame = r_str_Parame & "   FROM MNT_PARDES "
      If p_TipEmp = 1 Then
         r_str_Parame = r_str_Parame & "  WHERE PARDES_CODGRP = '507'"
      ElseIf p_TipEmp = 3 Then
         r_str_Parame = r_str_Parame & "  WHERE PARDES_CODGRP = '509'"
      End If
      r_str_Parame = r_str_Parame & "    AND PARDES_CODITE <> '000000' "
      r_str_Parame = r_str_Parame & "  ORDER BY PARDES_CODITE ASC "
      
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If r_rst_Genera.BOF And r_rst_Genera.EOF Then
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
         MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      grd_Listad.Redraw = False
      r_rst_Genera.MoveFirst
      Do While Not r_rst_Genera.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = r_rst_Genera!PARDES_CODITE
         
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(r_rst_Genera!PARDES_DESCRI)
         
         r_rst_Genera.MoveNext
      Loop
      grd_Listad.Redraw = True
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      
   ElseIf p_TipEmp = 2 Then
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT *  "
      r_str_Parame = r_str_Parame & "   FROM MNT_SEGEMP A "
      r_str_Parame = r_str_Parame & "  ORDER BY SEGEMP_CODIGO ASC "
      
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If r_rst_Genera.BOF And r_rst_Genera.EOF Then
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
         MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      grd_Listad.Redraw = False
      r_rst_Genera.MoveFirst
      Do While Not r_rst_Genera.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = r_rst_Genera!SEGEMP_CODIGO
         
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(r_rst_Genera!SEGEMP_RAZSOC)
         
         r_rst_Genera.MoveNext
      Loop
      grd_Listad.Redraw = True
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
   End If
   
   If p_TipEmp = 2 Then
      cmd_Agrega.Enabled = False
   End If
   
   If grd_Listad.Rows > 0 Then
      If p_TipEmp = 1 Then
         cmd_Agrega.Enabled = True
         cmd_Editar.Enabled = True
         cmd_Borrar.Enabled = True
         cmd_Perito.Enabled = True
         cmd_Imprim.Enabled = False
      End If
      If p_TipEmp = 2 Then
         cmd_Perito.Enabled = True
         cmd_Imprim.Enabled = False
      End If
      If p_TipEmp = 3 Then
         cmd_Perito.Enabled = True
         cmd_PerCon.Enabled = False
         cmd_Imprim.Enabled = False
      End If
      cmd_PerCon.Enabled = True
      
      'Ordenando por Razón Social
      pnl_Tit_RazSoc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
      
      grd_Listad.Enabled = True
      Call gs_UbiIniGrid(grd_Listad)
      Call gs_SetFocus(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_DblClick()
   If moddat_g_int_TipRec = 1 Then
      Call cmd_Editar_Click
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_CodEmp_Click()
   If Len(Trim(pnl_Tit_CodEmp.Tag)) = 0 Or pnl_Tit_CodEmp.Tag = "D" Then
      pnl_Tit_CodEmp.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_CodEmp.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_RazSoc_Click()
   If Len(Trim(pnl_Tit_RazSoc.Tag)) = 0 Or pnl_Tit_RazSoc.Tag = "D" Then
      pnl_Tit_RazSoc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_RazSoc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub
