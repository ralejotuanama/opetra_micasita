VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Tra_EvaSeg_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8460
   ClientLeft      =   1050
   ClientTop       =   2595
   ClientWidth     =   14115
   Icon            =   "OpeTra_frm_280.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   14115
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8475
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14115
      _Version        =   65536
      _ExtentX        =   24897
      _ExtentY        =   14949
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
         TabIndex        =   5
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
         Begin VB.CommandButton cmd_Export 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_280.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar datos a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_EvaSol 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_280.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13410
            Picture         =   "OpeTra_frm_280.frx":0BE0
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6015
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
         _ExtentY        =   10610
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
            Height          =   5955
            Left            =   30
            TabIndex        =   0
            Top             =   30
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   10504
            _Version        =   393216
            Rows            =   30
            Cols            =   14
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   7
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   690
            TabIndex        =   8
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito Hipotecario"
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
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   690
            TabIndex        =   9
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación de Seguros"
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
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_280.frx":1022
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   900
         Left            =   60
         TabIndex        =   10
         Top             =   7500
         Width           =   13995
         _Version        =   65536
         _ExtentX        =   24686
         _ExtentY        =   1587
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   240
            TabIndex        =   11
            Top             =   120
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "OPERACIONES NUEVAS"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   4500
            TabIndex        =   12
            Top             =   120
            Width           =   3405
            _Version        =   65536
            _ExtentX        =   6006
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SEGURO REGISTRADO - EN EVALUACION"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_OpeNue 
            Height          =   315
            Left            =   2520
            TabIndex        =   13
            Top             =   120
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_SegReg_Eva 
            Height          =   315
            Left            =   8100
            TabIndex        =   14
            Top             =   120
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_SegObs 
            Height          =   315
            Left            =   2520
            TabIndex        =   15
            Top             =   480
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_SegReg_Obs 
            Height          =   315
            Left            =   8100
            TabIndex        =   16
            Top             =   480
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
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
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   315
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SEGURO OBSERVADO"
            ForeColor       =   33023
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   315
            Left            =   4500
            TabIndex        =   18
            Top             =   480
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SEGURO REGISTRADO - OBSERVADO"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_TotReg 
            Height          =   315
            Left            =   12480
            TabIndex        =   19
            Top             =   480
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
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
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   315
            Left            =   10200
            TabIndex        =   20
            Top             =   480
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "TOTAL DE REGISTROS"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            FloodColor      =   0
            Font3D          =   2
            Alignment       =   1
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_EvaSeg_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_OpeNue        As Integer
Dim l_int_SegReg_Eva    As Integer
Dim l_int_SegReg_Obs    As Integer
Dim l_int_SegObs        As Integer

Private Sub cmd_EvaSol_Click()
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
   
   grd_Listad.Col = 9
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 10
   moddat_g_str_CodSub = grd_Listad.Text
   
   grd_Listad.Col = 13
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If moddat_g_int_TipMon <> 1 Then
      If moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon) = 0 Then
         MsgBox "No se encontró Tipo de Cambio registrado para " & moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon)) & ".", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   moddat_g_int_FlgAct = 1
   
   frm_Tra_EvaSeg_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      
      Call fs_Buscar
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Export_Click()
  'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub
Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_Conta      As Integer
Dim r_int_fila       As Integer

   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
      
   With r_obj_Excel.ActiveSheet
      r_int_NroFil = 2
      
      .Cells(1, 8) = "FECHA IMPRESION: " & Format(date, "dd/mm/yyyy")
      .Range(.Cells(1, 8), .Cells(1, 9)).Merge
      .Cells(1, 8).HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NroFil, 1) = "RELACION DE SOLICITUDES EN EVALUACIÓN DE SEGUROS"
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Range("A" & r_int_NroFil & ":I" & r_int_NroFil).Merge
      
      r_int_NroFil = 4
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 9)).Font.Bold = True
      
      .Columns("A").ColumnWidth = 30
      .Columns("B").ColumnWidth = 12
      .Columns("C").ColumnWidth = 9
      .Columns("D").ColumnWidth = 28
      .Columns("E").ColumnWidth = 9
      .Columns("F").ColumnWidth = 9
      .Columns("G").ColumnWidth = 14
      .Columns("H").ColumnWidth = 21
      .Columns("I").ColumnWidth = 14
      
      For r_int_Conta = r_int_NroFil To grd_Listad.Rows + 3
         .Cells(r_int_Conta, 1) = UCase(grd_Listad.TextMatrix(r_int_fila, 0))
         .Cells(r_int_Conta, 2) = UCase(grd_Listad.TextMatrix(r_int_fila, 1))
         .Cells(r_int_Conta, 3) = UCase(grd_Listad.TextMatrix(r_int_fila, 2))
         .Cells(r_int_Conta, 4) = UCase(grd_Listad.TextMatrix(r_int_fila, 3))
         .Cells(r_int_Conta, 5) = "'" & UCase(grd_Listad.TextMatrix(r_int_fila, 4))
         .Cells(r_int_Conta, 6) = "'" & UCase(grd_Listad.TextMatrix(r_int_fila, 5))
         .Cells(r_int_Conta, 7) = UCase(grd_Listad.TextMatrix(r_int_fila, 6))
         .Cells(r_int_Conta, 8) = UCase(grd_Listad.TextMatrix(r_int_fila, 7))
         .Cells(r_int_Conta, 9) = UCase(grd_Listad.TextMatrix(r_int_fila, 8))
         r_int_fila = r_int_fila + 1
      Next
            
      .Range(.Cells(2, 1), .Cells(2, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 1), .Cells(2, 9)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(2, 1), .Cells(2, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(2, 1), .Cells(2, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(2, 1), .Cells(2, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(2, 1), .Cells(2, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(2, 1), .Cells(2, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous

      .Range(.Cells(4, 1), .Cells(4, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 9)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 1), .Cells(4, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous

      .Range(.Cells(1, 1), .Cells(r_int_fila + 3, 9)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(r_int_fila + 3, 9)).Font.Size = 8
      .PageSetup.Orientation = xlLandscape
      
      .PageSetup.Zoom = 85
      .PageSetup.PrintArea = "A1:I" & r_int_fila + 3
   End With
   r_obj_Excel.Sheets(1).Name = "Solicitudes"
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
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
   grd_Listad.ColWidth(0) = 1895
   grd_Listad.ColWidth(1) = 1375
   grd_Listad.ColWidth(2) = 1235
   grd_Listad.ColWidth(3) = 3455
   grd_Listad.ColWidth(4) = 1195
   grd_Listad.ColWidth(5) = 1195
   grd_Listad.ColWidth(6) = 1670
   grd_Listad.ColWidth(7) = 2870
   grd_Listad.ColWidth(8) = 1580
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar()
   l_int_OpeNue = 0
   l_int_SegObs = 0
   l_int_SegReg_Eva = 0
   l_int_SegReg_Obs = 0
   
   'Obtener Tasa de Interes de Producto
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SOLMAE_CODPRD, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_FECSOL, "
   g_str_Parame = g_str_Parame & "       SEGUIM_FECINI, SEGUIM_SITUAC, SOLMAE_CONHIP, SOLMAE_CODSUB, SOLMAE_TIPMON, "
   g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT) ||' '|| Trim(DATGEN_APEMAT) ||' '|| Trim(DATGEN_NOMBRE) AS NOM_CLIENTE, "
   g_str_Parame = g_str_Parame & "       TRIM(PARDES_DESCRI) AS SITUACION,"
   g_str_Parame = g_str_Parame & "       (SELECT TRIM(PRODUC_DESCRI) FROM CRE_PRODUC WHERE PRODUC_CODIGO = LPAD(SOLMAE_CODPRD,3,'0')) AS NOM_PRODUCTO "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN TRA_SEGUIM B ON B.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND SEGUIM_CODINS = 42 AND (SEGUIM_SITUAC = 9 OR SEGUIM_SITUAC = 3) "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = A.SOLMAE_TITTDO AND C.DATGEN_NUMDOC = A.SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 23 AND D.PARDES_CODITE = SEGUIM_SITUAC "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND SOLMAE_CODINS = 41 "
   g_str_Parame = g_str_Parame & " ORDER BY SOLMAE_NUMERO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
     
   'CABECERA
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.FixedRows = 1

   grd_Listad.Row = 0
   grd_Listad.Col = 0:   grd_Listad.Text = "Producto":               grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 1:   grd_Listad.Text = "Nro. Solicitud":         grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 2:   grd_Listad.Text = "DOI Cliente":            grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 3:   grd_Listad.Text = "Apellidos y Nombres":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 4:   grd_Listad.Text = "F. Solicitud":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 5:   grd_Listad.Text = "F. Ing. Inst.":          grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 6:   grd_Listad.Text = "Situación Instancia":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 7:   grd_Listad.Text = "Situación Operación":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 8:   grd_Listad.Text = "Consej. Hipotecario":    grd_Listad.CellAlignment = flexAlignCenterCenter

   grd_Listad.Rows = grd_Listad.Rows - 1
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
         'g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' AND "
         'g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 42 AND "
         'g_str_Parame = g_str_Parame & "(SEGUIM_SITUAC = 9 OR SEGUIM_SITUAC = 3)"
         
         'If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         '   Exit Sub
         'End If
               
         'If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         '   g_rst_Genera.MoveFirst
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0
            grd_Listad.Text = g_rst_Princi!NOM_PRODUCTO 'moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
            
            grd_Listad.Col = 1
            grd_Listad.Text = Left(g_rst_Princi!SOLMAE_NUMERO, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Right(g_rst_Princi!SOLMAE_NUMERO, 4)
            
            grd_Listad.Col = 2
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
            
            grd_Listad.Col = 3
            grd_Listad.Text = g_rst_Princi!NOM_CLIENTE 'moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
            
            grd_Listad.Col = 4
            grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
            
            grd_Listad.Col = 5
            grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))

            grd_Listad.Col = 6
            grd_Listad.Text = g_rst_Princi!SITUACION 'moddat_gf_Consulta_ParDes("023", CStr(g_rst_Princi!SEGUIM_SITUAC))
                       
            grd_Listad.Col = 7
            grd_Listad.Text = fs_Buscar_SituacOper(g_rst_Princi!SOLMAE_NUMERO)
                       
            grd_Listad.Col = 8
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
            
            grd_Listad.Col = 9
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
            
            grd_Listad.Col = 10
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
            
            grd_Listad.Col = 11
            grd_Listad.Text = CStr(g_rst_Princi!SEGUIM_FECINI)
            
            grd_Listad.Col = 12
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_FECSOL)
         
            grd_Listad.Col = 13
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TIPMON)
                        
         'End If
         
         'g_rst_Genera.Close
         'Set g_rst_Genera = Nothing

         g_rst_Princi.MoveNext
      Loop
      
      'Ordenando por Nombre de Cliente
     
      Call gs_SorteaGrid(grd_Listad, 3, "C")
      
      pnl_OpeNue.Caption = l_int_OpeNue & " "
      pnl_SegReg_Eva.Caption = l_int_SegReg_Eva & " "
      pnl_SegReg_Obs.Caption = l_int_SegReg_Obs & " "
      pnl_SegObs.Caption = l_int_SegObs & " "
      pnl_TotReg.Caption = l_int_OpeNue + l_int_SegReg_Eva + l_int_SegReg_Obs + l_int_SegObs & " "
      
      grd_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad)
   Else
      cmd_EvaSol.Enabled = False
      
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Function fs_Buscar_SituacOper(ByVal p_NumSol As String) As String
Dim r_bol_FlgSeg     As Boolean

   r_bol_FlgSeg = False
   fs_Buscar_SituacOper = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SEGDET_CODOCU, (SELECT COUNT(*) FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & p_NumSol & "' AND SEGDET_CODINS = 42) AS CONTADOR"
   g_str_Parame = g_str_Parame & "  FROM TRA_SEGDET "
   g_str_Parame = g_str_Parame & " WHERE SEGDET_NUMSOL = '" & p_NumSol & "' "
   g_str_Parame = g_str_Parame & "   AND SEGDET_CODINS = 42 "
   g_str_Parame = g_str_Parame & " ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
   
   g_rst_GenAux.MoveFirst
   Do While Not g_rst_GenAux.EOF
      If g_rst_GenAux!CONTADOR = 1 Then
         r_bol_FlgSeg = True
         grd_Listad.Col = 7
         fs_Buscar_SituacOper = "OPERACION NUEVA"
         l_int_OpeNue = l_int_OpeNue + 1
         grd_Listad.CellForeColor = modgen_g_con_ColRoj
    
      ElseIf g_rst_GenAux!SEGDET_CODOCU = 35 Then
         r_bol_FlgSeg = True
         grd_Listad.Col = 6
         If grd_Listad.Text = "OBSERVADA" Then
            grd_Listad.Col = 7
            l_int_SegReg_Obs = l_int_SegReg_Obs + 1
            grd_Listad.CellForeColor = modgen_g_con_ColAzu
         Else
            grd_Listad.Col = 7
            l_int_SegReg_Eva = l_int_SegReg_Eva + 1
            grd_Listad.CellForeColor = modgen_g_con_ColVer
         End If
         fs_Buscar_SituacOper = "SEGURO REGISTRADO"
         g_rst_GenAux.Close
         Set g_rst_GenAux = Nothing
         Exit Function
      End If
      g_rst_GenAux.MoveNext
   Loop
   If r_bol_FlgSeg = False Then
      grd_Listad.Col = 6
      If grd_Listad.Text = "OBSERVADA" Then
         grd_Listad.Col = 7
         l_int_SegObs = l_int_SegObs + 1
         grd_Listad.CellForeColor = modgen_g_con_ColNar
         fs_Buscar_SituacOper = "SEGURO OBSERVADO"
      Else
         grd_Listad.Col = 7
         l_int_SegReg_Eva = l_int_SegReg_Eva + 1
         grd_Listad.CellForeColor = modgen_g_con_ColVer
         fs_Buscar_SituacOper = "SEGURO REGISTRADO"
      End If
   End If
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing

End Function

Private Sub grd_Listad_Click()
 Static Modo        As Boolean
   If grd_Listad.MouseRow = 0 Then
       If grd_Listad.MouseCol = 4 Then
          grd_Listad.Col = 12
       ElseIf grd_Listad.MouseCol = 5 Then
          grd_Listad.Col = 11
       Else
          grd_Listad.Col = grd_Listad.MouseCol
       End If
       ' Ordena en forma ascendente
       If Modo Then
           grd_Listad.Sort = 8 '2
           Modo = False
       ' Ordena en forma descendente
       Else
           grd_Listad.Sort = 7 '1
           Modo = True
       End If
       If grd_Listad.Rows > 1 Then
          Call gs_UbicaGrid(grd_Listad, 1)
       Else
          Call gs_UbicaGrid(grd_Listad, 0)
       End If
   End If
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_EvaSol_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

