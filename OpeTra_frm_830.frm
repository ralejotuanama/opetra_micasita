VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Ges_TecPro_08 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   Icon            =   "OpeTra_frm_830.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8685
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   7935
      _Version        =   65536
      _ExtentX        =   13996
      _ExtentY        =   15319
      _StockProps     =   15
      BackColor       =   14215660
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
         Width           =   7845
         _Version        =   65536
         _ExtentX        =   13838
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
            TabIndex        =   13
            Top             =   60
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Entidades Técnicas"
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
            Left            =   660
            TabIndex        =   14
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Mantenimiento de Ventas y Patrimonio"
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
            Left            =   90
            Picture         =   "OpeTra_frm_830.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   740
         Width           =   7845
         _Version        =   65536
         _ExtentX        =   13838
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_830.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Borrar Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_830.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Nueva Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_830.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Modificar Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7230
            Picture         =   "OpeTra_frm_830.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Ventana"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5595
         Left            =   30
         TabIndex        =   7
         Top             =   3030
         Width           =   7845
         _Version        =   65536
         _ExtentX        =   13838
         _ExtentY        =   9869
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
            Height          =   5175
            Left            =   30
            TabIndex        =   8
            Top             =   360
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   9128
            _Version        =   393216
            Rows            =   24
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_CodAno 
            Height          =   285
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   705
            _Version        =   65536
            _ExtentX        =   1244
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
         Begin Threed.SSPanel pnl_Tit_CodMes 
            Height          =   285
            Left            =   750
            TabIndex        =   10
            Top             =   60
            Width           =   675
            _Version        =   65536
            _ExtentX        =   1191
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   1410
            TabIndex        =   11
            Top             =   60
            Width           =   2050
            _Version        =   65536
            _ExtentX        =   3616
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Ventas"
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
            Left            =   3450
            TabIndex        =   12
            Top             =   60
            Width           =   2020
            _Version        =   65536
            _ExtentX        =   3563
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Patrimonio"
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
            Left            =   5460
            TabIndex        =   24
            Top             =   60
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Capital Social"
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   1575
         Left            =   30
         TabIndex        =   15
         Top             =   1410
         Width           =   7845
         _Version        =   65536
         _ExtentX        =   13838
         _ExtentY        =   2778
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
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   16
            Top             =   840
            Width           =   5625
            _Version        =   65536
            _ExtentX        =   9922
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
         Begin Threed.SSPanel pnl_NroDoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   17
            Top             =   480
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
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
         Begin Threed.SSPanel pnl_TipEmp 
            Height          =   315
            Left            =   1620
            TabIndex        =   18
            Top             =   1200
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
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
         Begin Threed.SSPanel pnl_TipDoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   22
            Top             =   120
            Width           =   5625
            _Version        =   65536
            _ExtentX        =   9922
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
         Begin VB.Label lbl_TipDoc 
            Caption         =   "Tipo Documento:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   150
            Width           =   1335
         End
         Begin VB.Label lbl_NumDoc 
            Caption         =   "Nro. Documento:"
            Height          =   225
            Left            =   120
            TabIndex        =   21
            Top             =   525
            Width           =   1335
         End
         Begin VB.Label lbl_RazSoc 
            Caption         =   "Razón Social:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   870
            Width           =   1335
         End
         Begin VB.Label lbl_TipEmp 
            Caption         =   "Tipo Empresa:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1230
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_FlgAct = 1
   
   frm_Ges_TecPro_09.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_CodAno = grd_Listad.Text
   
   grd_Listad.Col = 1
   moddat_g_str_CodMes = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   
   frm_Ges_TecPro_09.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_CodAno = grd_Listad.Text
   
   grd_Listad.Col = 1
   moddat_g_str_CodMes = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM CTB_VTAPAT "
   g_str_Parame = g_str_Parame & " WHERE VTAPAT_TIPDOC = " & moddat_g_int_TipDoc
   g_str_Parame = g_str_Parame & "   AND VTAPAT_NUMDOC = '" & moddat_g_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "   AND VTAPAT_CODANO = " & moddat_g_str_CodAno
   g_str_Parame = g_str_Parame & "   AND VTAPAT_CODMES = " & moddat_g_str_CodMes
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
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
   grd_Listad.ColWidth(0) = 700
   grd_Listad.ColWidth(1) = 660
   grd_Listad.ColWidth(2) = 2020
   grd_Listad.ColWidth(3) = 2010
   grd_Listad.ColWidth(4) = 1990
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
   
   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", moddat_g_int_TipDoc)
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT VTAPAT_CODANO, VTAPAT_CODMES, VTAPAT_MTOVTA, VTAPAT_MTOPAT, VTAPAT_MTOCAP "
   g_str_Parame = g_str_Parame & "   FROM CTB_VTAPAT "
   g_str_Parame = g_str_Parame & "  WHERE VTAPAT_TIPDOC = " & moddat_g_int_TipDoc & ""
   g_str_Parame = g_str_Parame & "    AND VTAPAT_NUMDOC = '" & moddat_g_str_NumDoc & "'"
   g_str_Parame = g_str_Parame & "  ORDER BY VTAPAT_CODANO, VTAPAT_CODMES ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Trim(g_rst_Princi!VTAPAT_CODANO)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Format(Trim(g_rst_Princi!VTAPAT_CODMES), "00")
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(Trim(g_rst_Princi!VTAPAT_MTOVTA), "###,###,###,##0.00")
      
      grd_Listad.Col = 3
      grd_Listad.Text = Format(Trim(g_rst_Princi!VTAPAT_MTOPAT), "###,###,###,##0.00")
      
      grd_Listad.Col = 4
      grd_Listad.Text = Format(Trim(g_rst_Princi!VTAPAT_MTOCAP), "###,###,###,##0.00")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

