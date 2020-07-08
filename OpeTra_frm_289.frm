VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Pro_EvaPBP_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6495
   ClientLeft      =   5010
   ClientTop       =   3465
   ClientWidth     =   12870
   Icon            =   "OpeTra_frm_289.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12885
      _Version        =   65536
      _ExtentX        =   22728
      _ExtentY        =   11456
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
         Height          =   4995
         Left            =   30
         TabIndex        =   5
         Top             =   1440
         Width           =   12795
         _Version        =   65536
         _ExtentX        =   22569
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Período"
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
            Left            =   2040
            TabIndex        =   7
            Top             =   60
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Evaluados"
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
            Left            =   4110
            TabIndex        =   8
            Top             =   60
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Premios Asignados"
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
            Left            =   6180
            TabIndex        =   9
            Top             =   60
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Premios Perdidos"
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
            Left            =   10320
            TabIndex        =   10
            Top             =   60
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   4605
            Left            =   30
            TabIndex        =   11
            Top             =   360
            Width           =   12705
            _ExtentX        =   22410
            _ExtentY        =   8123
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   8250
            TabIndex        =   14
            Top             =   60
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Premios Pendientes"
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
         TabIndex        =   1
         Top             =   30
         Width           =   12795
         _Version        =   65536
         _ExtentX        =   22569
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
            Height          =   555
            Left            =   570
            TabIndex        =   2
            Top             =   30
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Evaluación y Asignación de Premio Buen Pagador"
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
            Picture         =   "OpeTra_frm_289.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   3
         Top             =   750
         Width           =   12795
         _Version        =   65536
         _ExtentX        =   22569
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
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_289.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Editar Evaluación PBP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NuePro 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_289.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Nueva Evaluación PBP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12180
            Picture         =   "OpeTra_frm_289.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_EvaPBP_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Editar_Click()
   Dim r_int_Situac     As Integer

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 6
   moddat_g_str_Codigo = grd_Listad.Text
   
   grd_Listad.Col = 7
   moddat_g_str_CodIte = grd_Listad.Text
   
   grd_Listad.Col = 8
   moddat_g_int_Situac = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If moddat_g_int_Situac <> 1 Then
      MsgBox "El Período ya fue asignado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Pro_EvaPBP_03.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_NuePro_Click()
   moddat_g_int_FlgAct = 1

   frm_Pro_EvaPBP_02.Show 1
   
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
   
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1985
   grd_Listad.ColWidth(1) = 2075
   grd_Listad.ColWidth(2) = 2075
   grd_Listad.ColWidth(3) = 2075
   grd_Listad.ColWidth(4) = 2075
   grd_Listad.ColWidth(5) = 2075
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Buscar()
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = "SELECT * FROM CRE_CABPBP "
   g_str_Parame = g_str_Parame & "ORDER BY CABPBP_PERANO DESC, CABPBP_PERMES DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "No se pudo leer la tabla CBR_REGACC.", vbCritical, modgen_g_str_NomPlt
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
      grd_Listad.Text = moddat_gf_Consulta_ParDes("033", CStr(g_rst_Princi!CABPBP_PERMES)) & " - " & Format(g_rst_Princi!CABPBP_PERANO, "0000")
      
      grd_Listad.Col = 1
      grd_Listad.Text = Format(g_rst_Princi!CABPBP_TOTEVA, "###,##0")
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!CABPBP_TOTASG, "###,##0")
      
      grd_Listad.Col = 3
      grd_Listad.Text = Format(g_rst_Princi!CABPBP_TOTPER, "###,##0")
      
      grd_Listad.Col = 4
      grd_Listad.Text = Format(g_rst_Princi!CABPBP_TOTEVA - g_rst_Princi!CABPBP_TOTASG - g_rst_Princi!CABPBP_TOTPER, "###,##0")
      
      grd_Listad.Col = 5
      grd_Listad.Text = moddat_gf_Consulta_ParDes("274", CStr(g_rst_Princi!CABPBP_SITUAC))
      
      grd_Listad.Col = 6
      grd_Listad.Text = CStr(g_rst_Princi!CABPBP_PERMES)
      
      grd_Listad.Col = 7
      grd_Listad.Text = CStr(g_rst_Princi!CABPBP_PERANO)
      
      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!CABPBP_SITUAC)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

