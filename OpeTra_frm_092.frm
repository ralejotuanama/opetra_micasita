VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_EmpPer_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6540
   ClientLeft      =   2640
   ClientTop       =   1740
   ClientWidth     =   8205
   Icon            =   "OpeTra_frm_092.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   11562
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
         Height          =   525
         Left            =   30
         TabIndex        =   1
         Top             =   1440
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   926
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
         Begin Threed.SSPanel pnl_EmpPer 
            Height          =   405
            Left            =   1440
            TabIndex        =   2
            Top             =   60
            Width           =   6585
            _Version        =   65536
            _ExtentX        =   11615
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "SSPanel3"
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
         Begin VB.Label lbl_NomEmp 
            Caption         =   "Empresa Peritaje:"
            Height          =   195
            Left            =   60
            TabIndex        =   3
            Top             =   150
            Width           =   1350
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   30
         TabIndex        =   4
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7500
            Picture         =   "OpeTra_frm_092.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_092.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imprimir Listado"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_092.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_092.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_092.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   30
         TabIndex        =   10
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
            Height          =   270
            Left            =   630
            TabIndex        =   11
            Top             =   60
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Empresas de Peritaje"
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
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   270
            Left            =   630
            TabIndex        =   12
            Top             =   330
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Personas de Contacto"
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
            Left            =   7680
            Top             =   30
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentaci�n Preliminar"
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
            Picture         =   "OpeTra_frm_092.frx":11AE
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   4485
         Left            =   30
         TabIndex        =   13
         Top             =   2010
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
         _ExtentY        =   7911
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
         Begin Threed.SSPanel pnl_Tit_NomCon 
            Height          =   285
            Left            =   1500
            TabIndex        =   14
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Perito"
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
         Begin Threed.SSPanel pnl_Tit_CodCon 
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "C�digo Perito"
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
            Height          =   4095
            Left            =   30
            TabIndex        =   16
            Top             =   360
            Width           =   8025
            _ExtentX        =   14155
            _ExtentY        =   7223
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
   End
End
Attribute VB_Name = "frm_EmpPer_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgAct_1 = 1
   moddat_g_int_FlgGrb = 1
   
   frm_EmpPer_06.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("�Est� seguro de eliminar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Obteniendo Informaci�n del Registro
   g_str_Parame = "USP_MNT_PERCON_BORRAR ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
   
   'Tipo de empresa
   If moddat_g_int_TipRec = 1 Then
      g_str_Parame = g_str_Parame & "1) "
   ElseIf moddat_g_int_TipRec = 2 Then
      g_str_Parame = g_str_Parame & "2) "
   ElseIf moddat_g_int_TipRec = 3 Then
      g_str_Parame = g_str_Parame & "3) "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
   
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgAct_1 = 1
   moddat_g_int_FlgGrb = 2
   
   frm_EmpPer_06.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Imprim_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   If MsgBox("�Est� seguro de Imprimir el reporte?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "MNT_PARDES"
   crp_Imprim.DataFiles(1) = "MNT_PERCON"
   
   'Tipo de empresa
   If moddat_g_int_TipRec = 1 Then
      crp_Imprim.SelectionFormula = "{MNT_PARDES.PARDES_CODGRP} = '507' AND {MNT_PERCON.PERCON_CODEMP} = '" & moddat_g_str_CodGrp & "' AND {MNT_PERCON.PERCON_TIPTAB} = 1 "
   ElseIf moddat_g_int_TipRec = 2 Then
      crp_Imprim.SelectionFormula = "{MNT_PARDES.PARDES_CODGRP} = '507' AND {MNT_PERCON.PERCON_CODEMP} = '" & moddat_g_str_CodGrp & "' AND {MNT_PERCON.PERCON_TIPTAB} = 2 "
   ElseIf moddat_g_int_TipRec = 3 Then
      crp_Imprim.SelectionFormula = "{MNT_PARDES.PARDES_CODGRP} = '509' AND {MNT_PERCON.PERCON_CODEMP} = '" & moddat_g_str_CodGrp & "' AND {MNT_PERCON.PERCON_TIPTAB} = 3 "
   End If
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "MNT_EMPPER_03.RPT"
      
   'Tipo de empresa
   If moddat_g_int_TipRec = 1 Then
      crp_Imprim.ParameterFields(0) = "p_Titulo;" & "REPORTE DE CONTACTOS POR EMPRESA DE PERITAJE" & ";True"
   ElseIf moddat_g_int_TipRec = 2 Then
      crp_Imprim.ParameterFields(0) = "p_Titulo;" & "REPORTE DE CONTACTOS POR EMPRESA DE SEGUROS" & ";True"
   ElseIf moddat_g_int_TipRec = 3 Then
      crp_Imprim.ParameterFields(0) = "p_Titulo;" & "REPORTE DE CONTACTOS POR EMPRESA DE NOTARIA" & ";True"
   End If
   
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call gs_CentraForm(Me)
   
   Me.Caption = modgen_g_str_NomPlt
   pnl_EmpPer.Caption = moddat_g_str_CodGrp & " - " & moddat_g_str_DesGrp
   
   Call fs_Inicia
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1455:      grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColWidth(1) = 6155:      grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   'Tipo de empresa
   If moddat_g_int_TipRec = 1 Then
      pnl_TitPri.Caption = "Empresas de Peritaje"
      lbl_NomEmp.Caption = "Empresa Peritaje:"
   ElseIf moddat_g_int_TipRec = 2 Then
      pnl_TitPri.Caption = "Empresas de Seguros"
      lbl_NomEmp.Caption = "Empresa Seguros:"
   ElseIf moddat_g_int_TipRec = 3 Then
      pnl_TitPri.Caption = "Empresas de Notaria"
      lbl_NomEmp.Caption = "Notaria:"
   End If
End Sub

Private Sub fs_Buscar()
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   'cmd_Imprim.Enabled = False
   
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * "
   g_str_Parame = g_str_Parame & "   FROM MNT_PERCON "
   g_str_Parame = g_str_Parame & "  WHERE PERCON_CODEMP = '" & moddat_g_str_CodGrp & "' "
   'Tipo de empresa
   If moddat_g_int_TipRec = 1 Then
      g_str_Parame = g_str_Parame & " AND PERCON_TIPTAB =  1 "
   ElseIf moddat_g_int_TipRec = 2 Then
      g_str_Parame = g_str_Parame & " AND PERCON_TIPTAB =  2 "
   ElseIf moddat_g_int_TipRec = 3 Then
      g_str_Parame = g_str_Parame & " AND PERCON_TIPTAB =  3 "
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY PERCON_NOMBRE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = g_rst_Genera!PERCON_CODCON
      
      grd_Listad.Col = 1
      grd_Listad.Text = g_rst_Genera!PERCON_NOMBRE
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      
      'If moddat_g_int_TipRec = 1 Then
      'S   cmd_Imprim.Enabled = True
      'End If
      
      'Ordenando por Raz�n Social
      pnl_Tit_NomCon.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
      
      grd_Listad.Enabled = True
      
      Call gs_UbiIniGrid(grd_Listad)
      Call gs_SetFocus(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_CodCon_Click()
   If Len(Trim(pnl_Tit_CodCon.Tag)) = 0 Or pnl_Tit_CodCon.Tag = "D" Then
      pnl_Tit_CodCon.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_CodCon.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCon_Click()
   If Len(Trim(pnl_Tit_NomCon.Tag)) = 0 Or pnl_Tit_NomCon.Tag = "D" Then
      pnl_Tit_NomCon.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NomCon.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

