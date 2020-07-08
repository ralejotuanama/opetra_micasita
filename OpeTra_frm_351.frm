VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_Rpt_ClaCar_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14280
   Icon            =   "OpeTra_frm_351.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   14280
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8040
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14280
      _Version        =   65536
      _ExtentX        =   25188
      _ExtentY        =   14182
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
         Left            =   60
         TabIndex        =   6
         Top             =   780
         Width           =   14160
         _Version        =   65536
         _ExtentX        =   24977
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_351.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_351.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13560
            Picture         =   "OpeTra_frm_351.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   885
         Left            =   60
         TabIndex        =   7
         Top             =   1470
         Width           =   14160
         _Version        =   65536
         _ExtentX        =   24977
         _ExtentY        =   1561
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   3765
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1140
            TabIndex        =   1
            Top             =   480
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   195
            Left            =   270
            TabIndex        =   12
            Top             =   510
            Width           =   825
         End
         Begin VB.Label Label5 
            Caption         =   "Periodo:"
            Height          =   225
            Left            =   270
            TabIndex        =   11
            Top             =   150
            Width           =   825
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   14160
         _Version        =   65536
         _ExtentX        =   24977
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
            Left            =   600
            TabIndex        =   9
            Top             =   60
            Width           =   8835
            _Version        =   65536
            _ExtentX        =   15584
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Reporte de Clasificación de Cartera"
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
            Picture         =   "OpeTra_frm_351.frx":0A62
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   5595
         Left            =   60
         TabIndex        =   10
         Top             =   2400
         Width           =   14160
         _Version        =   65536
         _ExtentX        =   24977
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
         Begin TabDlg.SSTab tab_Clasif 
            Height          =   5460
            Left            =   75
            TabIndex        =   13
            Top             =   75
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   9631
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Clasificacion"
            TabPicture(0)   =   "OpeTra_frm_351.frx":0D6C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_LisCalif"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Detalle por Clasificacion"
            TabPicture(1)   =   "OpeTra_frm_351.frx":0D88
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_LisDetCalif"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Detalle por Producto"
            TabPicture(2)   =   "OpeTra_frm_351.frx":0DA4
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_LisDetProd"
            Tab(2).ControlCount=   1
            Begin MSFlexGridLib.MSFlexGrid grd_LisCalif 
               Height          =   5055
               Left            =   60
               TabIndex        =   14
               Top             =   345
               Width           =   13935
               _ExtentX        =   24580
               _ExtentY        =   8916
               _Version        =   393216
               Rows            =   0
               Cols            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisDetCalif 
               Height          =   5055
               Left            =   -74940
               TabIndex        =   15
               Top             =   345
               Width           =   13935
               _ExtentX        =   24580
               _ExtentY        =   8916
               _Version        =   393216
               Rows            =   0
               Cols            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisDetProd 
               Height          =   5055
               Left            =   -74940
               TabIndex        =   16
               Top             =   345
               Width           =   13935
               _ExtentX        =   24580
               _ExtentY        =   8916
               _Version        =   393216
               Rows            =   0
               Cols            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_ClaCar_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_int_PerMes     As Integer
Dim r_int_PerAno     As Integer
Dim r_obj_Excel      As Excel.Application

Private Sub fs_GenExc_Clasificacion()
   Dim r_int_Contad     As Integer
   
   With r_obj_Excel.Sheets(1)
      'Titulo
      .Name = "CLASIF."
      .Cells(2, 2) = "CLASIFICACION DE CARTERA DEL MES " & UCase(Trim(cmb_PerMes.Text)) & " DEL " & CStr(ipp_PerAno.Text)
      .Range(.Cells(2, 2), .Cells(2, 13)).Merge
      .Range("B2:M2").HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 2), .Cells(2, 13)).Font.Name = "Calibri"
      .Range(.Cells(2, 2), .Cells(2, 13)).Font.Size = 12
      .Range(.Cells(2, 2), .Cells(2, 13)).Font.Bold = True
               
      .Columns("C").NumberFormat = "###,##0"
      .Columns("D").NumberFormat = "###,##0"
      .Columns("E").NumberFormat = "###,##0"
      .Columns("F").NumberFormat = "###,##0"
      .Columns("G").NumberFormat = "###,##0"
      .Columns("H").NumberFormat = "###,##0"
      .Columns("I").NumberFormat = "###,##0"
      .Columns("J").NumberFormat = "###,##0"
      .Columns("K").NumberFormat = "###,##0"
      .Columns("L").NumberFormat = "###,##0"
      .Columns("M").NumberFormat = "###,##0"
      
      .Columns("A").ColumnWidth = 4
      .Columns("B").ColumnWidth = 20
      .Columns("B").ColumnWidth = 20
      .Columns("D").ColumnWidth = 15
      .Columns("F").ColumnWidth = 15
      .Columns("H").ColumnWidth = 15
      .Columns("J").ColumnWidth = 15
      .Columns("L").ColumnWidth = 15
      .Columns("M").ColumnWidth = 15
      
      .Range(.Cells(4, 2), .Cells(7, 2)).Merge
      .Range(.Cells(4, 2), .Cells(7, 2)) = "Producto"
      .Range(.Cells(4, 2), .Cells(7, 2)).WrapText = True
      .Range(.Cells(4, 2), .Cells(7, 2)).VerticalAlignment = xlCenter
      
      .Range(.Cells(4, 3), .Cells(4, 12)) = "Calificacion"
      .Range(.Cells(4, 3), .Cells(4, 12)).Merge
      .Range(.Cells(5, 3), .Cells(5, 4)) = "Normal"
      .Range(.Cells(5, 3), .Cells(5, 4)).Merge
      .Range(.Cells(5, 5), .Cells(5, 6)) = "CPP"
      .Range(.Cells(5, 5), .Cells(5, 6)).Merge
      .Range(.Cells(5, 7), .Cells(5, 8)) = "Deficiente"
      .Range(.Cells(5, 7), .Cells(5, 8)).Merge
      .Range(.Cells(5, 9), .Cells(5, 10)) = "Dudoso"
      .Range(.Cells(5, 9), .Cells(5, 10)).Merge
      .Range(.Cells(5, 11), .Cells(5, 12)) = "Perdida"
      .Range(.Cells(5, 11), .Cells(5, 12)).Merge
      .Range(.Cells(4, 13), .Cells(7, 13)).Merge
      .Range(.Cells(4, 13), .Cells(7, 13)) = "Total"
      .Range(.Cells(4, 13), .Cells(7, 13)).VerticalAlignment = xlCenter
      .Range(.Cells(6, 3), .Cells(6, 4)) = "01 - 30 dias"
      .Range(.Cells(6, 3), .Cells(6, 4)).Merge
      .Range(.Cells(6, 5), .Cells(6, 6)) = "31 - 60 dias"
      .Range(.Cells(6, 5), .Cells(6, 6)).Merge
      .Range(.Cells(6, 7), .Cells(6, 8)) = "61 - 120 dias"
      .Range(.Cells(6, 7), .Cells(6, 8)).Merge
      .Range(.Cells(6, 9), .Cells(6, 10)) = "121 - 365 dias"
      .Range(.Cells(6, 9), .Cells(6, 10)).Merge
      .Range(.Cells(6, 11), .Cells(6, 12)) = "mas de 365 dias"
      .Range(.Cells(6, 11), .Cells(6, 12)).Merge
      
      'r_obj_Excel.Visible = True

      For r_int_Contad = 0 To 8
         .Range(.Cells(7, r_int_Contad + 3), .Cells(7, r_int_Contad + 3)) = "N° Creditos"
         .Range(.Cells(7, r_int_Contad + 4), .Cells(7, r_int_Contad + 4)) = "Saldo"
         r_int_Contad = r_int_Contad + 1
      Next
      
      .Range(.Cells(4, 2), .Cells(7, 13)).HorizontalAlignment = xlVAlignCenter
      .Range(.Cells(4, 2), .Cells(7, 13)).Font.Name = "Calibri"
      .Range(.Cells(4, 2), .Cells(7, 13)).Font.Size = 10
      .Range(.Cells(4, 2), .Cells(7, 13)).Font.Bold = True
      
      .Range(.Cells(4, 2), .Cells(7, 13)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(7, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(5, 3), .Cells(5, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(6, 3), .Cells(6, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(7, 3), .Cells(7, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(8, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(7, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(7, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(7, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(8, 2), .Cells(9, 2)).Merge
      .Range(.Cells(8, 2), .Cells(9, 2)) = "CRC-PBP"
      .Range(.Cells(8, 2), .Cells(9, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(10, 2), .Cells(11, 2)).Merge
      .Range(.Cells(10, 2), .Cells(11, 2)) = "Micasita"
      .Range(.Cells(10, 2), .Cells(11, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(12, 2), .Cells(13, 2)).Merge
      .Range(.Cells(12, 2), .Cells(13, 2)) = "CME"
      .Range(.Cells(12, 2), .Cells(13, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(14, 2), .Cells(15, 2)).Merge
      .Range(.Cells(14, 2), .Cells(15, 2)) = "N. MiVivienda"
      .Range(.Cells(14, 2), .Cells(15, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(16, 2), .Cells(17, 2)).Merge
      .Range(.Cells(16, 2), .Cells(17, 2)) = "MiCasa Mas"
      .Range(.Cells(16, 2), .Cells(17, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(18, 2), .Cells(19, 2)).Merge
      .Range(.Cells(18, 2), .Cells(19, 2)) = "BBP"
      .Range(.Cells(18, 2), .Cells(19, 2)).VerticalAlignment = xlCenter
      
      .Range(.Cells(20, 2), .Cells(21, 2)).Merge
      .Range(.Cells(20, 2), .Cells(21, 2)) = "TECHO PROPIO"
      .Range(.Cells(20, 2), .Cells(21, 2)).VerticalAlignment = xlCenter
      
      
      .Range(.Cells(22, 2), .Cells(22, 2)).Merge
      .Range(.Cells(22, 2), .Cells(22, 2)) = "Promedio Ponderado"
      .Range(.Cells(22, 2), .Cells(22, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(24, 2), .Cells(24, 2)) = "Total"
      .Range(.Cells(24, 2), .Cells(24, 2)).VerticalAlignment = xlCenter
      
      'CRC-PBP
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO "
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrCRC & ") "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV "
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .Range(.Cells(8, 3), .Cells(8, 3)) = g_rst_Princi!CONT
            .Range(.Cells(8, 3), .Cells(8, 3)).VerticalAlignment = xlCenter
            .Range(.Cells(8, 4), .Cells(8, 4)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(8, 4), .Cells(8, 4)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .Range(.Cells(8, 5), .Cells(8, 5)) = g_rst_Princi!CONT
            .Range(.Cells(8, 5), .Cells(8, 5)).VerticalAlignment = xlCenter
            .Range(.Cells(8, 6), .Cells(8, 6)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(8, 6), .Cells(8, 6)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .Range(.Cells(8, 7), .Cells(8, 7)) = g_rst_Princi!CONT
            .Range(.Cells(8, 7), .Cells(8, 7)).VerticalAlignment = xlCenter
            .Range(.Cells(8, 8), .Cells(8, 8)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(8, 8), .Cells(8, 8)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .Range(.Cells(8, 9), .Cells(8, 9)) = g_rst_Princi!CONT
            .Range(.Cells(8, 9), .Cells(8, 9)).VerticalAlignment = xlCenter
            .Range(.Cells(8, 10), .Cells(8, 10)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(8, 10), .Cells(8, 10)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .Range(.Cells(8, 11), .Cells(8, 11)) = g_rst_Princi!CONT
            .Range(.Cells(8, 11), .Cells(8, 11)).VerticalAlignment = xlCenter
            .Range(.Cells(8, 12), .Cells(8, 12)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(8, 12), .Cells(8, 12)).VerticalAlignment = xlCenter
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'MICASITA
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .Range(.Cells(10, 3), .Cells(10, 3)) = g_rst_Princi!CONT
            .Range(.Cells(10, 3), .Cells(10, 3)).VerticalAlignment = xlCenter
            .Range(.Cells(10, 4), .Cells(10, 4)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(10, 4), .Cells(10, 4)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .Range(.Cells(10, 5), .Cells(10, 5)) = g_rst_Princi!CONT
            .Range(.Cells(10, 5), .Cells(10, 5)).VerticalAlignment = xlCenter
            .Range(.Cells(10, 6), .Cells(10, 6)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(10, 6), .Cells(10, 6)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .Range(.Cells(10, 7), .Cells(10, 7)) = g_rst_Princi!CONT
            .Range(.Cells(10, 7), .Cells(10, 7)).VerticalAlignment = xlCenter
            .Range(.Cells(10, 8), .Cells(10, 8)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(10, 8), .Cells(10, 8)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .Range(.Cells(10, 9), .Cells(10, 9)) = g_rst_Princi!CONT
            .Range(.Cells(10, 9), .Cells(10, 9)).VerticalAlignment = xlCenter
            .Range(.Cells(10, 10), .Cells(10, 10)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(10, 10), .Cells(10, 10)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .Range(.Cells(10, 11), .Cells(10, 11)) = g_rst_Princi!CONT
            .Range(.Cells(10, 11), .Cells(10, 11)).VerticalAlignment = xlCenter
            .Range(.Cells(10, 12), .Cells(10, 12)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(10, 12), .Cells(10, 12)).VerticalAlignment = xlCenter
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'CME
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrCME & ") "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .Range(.Cells(12, 3), .Cells(12, 3)) = g_rst_Princi!CONT
            .Range(.Cells(12, 3), .Cells(12, 3)).VerticalAlignment = xlCenter
            .Range(.Cells(12, 4), .Cells(12, 4)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(12, 4), .Cells(12, 4)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .Range(.Cells(12, 5), .Cells(12, 5)) = g_rst_Princi!CONT
            .Range(.Cells(12, 5), .Cells(12, 5)).VerticalAlignment = xlCenter
            .Range(.Cells(12, 6), .Cells(12, 6)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(12, 6), .Cells(12, 6)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .Range(.Cells(12, 7), .Cells(12, 7)) = g_rst_Princi!CONT
            .Range(.Cells(12, 7), .Cells(12, 7)).VerticalAlignment = xlCenter
            .Range(.Cells(12, 8), .Cells(12, 8)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(12, 8), .Cells(12, 8)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .Range(.Cells(12, 9), .Cells(12, 9)) = g_rst_Princi!CONT
            .Range(.Cells(12, 9), .Cells(12, 9)).VerticalAlignment = xlCenter
            .Range(.Cells(12, 10), .Cells(12, 10)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(12, 10), .Cells(12, 10)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .Range(.Cells(12, 11), .Cells(12, 11)) = g_rst_Princi!CONT
            .Range(.Cells(12, 11), .Cells(12, 11)).VerticalAlignment = xlCenter
            .Range(.Cells(12, 12), .Cells(12, 12)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(12, 12), .Cells(12, 12)).VerticalAlignment = xlCenter
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'MIVIVIENDA
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND (HIPCIE_CODPRD IN (" & moddat_g_str_AgrMIHG & "," & moddat_g_str_Agr2FMV & ") OR HIPCIE_CODPRD = '025') "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV "
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .Range(.Cells(14, 3), .Cells(14, 3)) = g_rst_Princi!CONT
            .Range(.Cells(14, 3), .Cells(14, 3)).VerticalAlignment = xlCenter
            .Range(.Cells(14, 4), .Cells(14, 4)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(14, 4), .Cells(14, 4)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .Range(.Cells(14, 5), .Cells(14, 5)) = g_rst_Princi!CONT
            .Range(.Cells(14, 5), .Cells(14, 5)).VerticalAlignment = xlCenter
            .Range(.Cells(14, 6), .Cells(14, 6)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(14, 6), .Cells(14, 6)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .Range(.Cells(14, 7), .Cells(14, 7)) = g_rst_Princi!CONT
            .Range(.Cells(14, 7), .Cells(14, 7)).VerticalAlignment = xlCenter
            .Range(.Cells(14, 8), .Cells(14, 8)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(14, 8), .Cells(14, 8)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .Range(.Cells(14, 9), .Cells(14, 9)) = g_rst_Princi!CONT
            .Range(.Cells(14, 9), .Cells(14, 9)).VerticalAlignment = xlCenter
            .Range(.Cells(14, 10), .Cells(14, 10)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(14, 10), .Cells(14, 10)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .Range(.Cells(14, 11), .Cells(14, 11)) = g_rst_Princi!CONT
            .Range(.Cells(14, 11), .Cells(14, 11)).VerticalAlignment = xlCenter
            .Range(.Cells(14, 12), .Cells(14, 12)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(14, 12), .Cells(14, 12)).VerticalAlignment = xlCenter
         End If
                    
         g_rst_Princi.MoveNext
      Loop

      'MICASAMAS
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD = '019'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .Range(.Cells(16, 3), .Cells(16, 3)) = g_rst_Princi!CONT
            .Range(.Cells(16, 3), .Cells(16, 3)).VerticalAlignment = xlCenter
            .Range(.Cells(16, 4), .Cells(16, 4)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(16, 4), .Cells(16, 4)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .Range(.Cells(16, 5), .Cells(16, 5)) = g_rst_Princi!CONT
            .Range(.Cells(16, 5), .Cells(16, 5)).VerticalAlignment = xlCenter
            .Range(.Cells(16, 6), .Cells(16, 6)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(16, 6), .Cells(16, 6)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .Range(.Cells(16, 7), .Cells(16, 7)) = g_rst_Princi!CONT
            .Range(.Cells(16, 7), .Cells(16, 7)).VerticalAlignment = xlCenter
            .Range(.Cells(16, 8), .Cells(16, 8)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(16, 8), .Cells(16, 8)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .Range(.Cells(16, 9), .Cells(16, 9)) = g_rst_Princi!CONT
            .Range(.Cells(16, 9), .Cells(16, 9)).VerticalAlignment = xlCenter
            .Range(.Cells(16, 10), .Cells(16, 10)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(16, 10), .Cells(16, 10)).VerticalAlignment = xlCenter
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .Range(.Cells(16, 11), .Cells(16, 11)) = g_rst_Princi!CONT
            .Range(.Cells(16, 11), .Cells(16, 11)).VerticalAlignment = xlCenter
            .Range(.Cells(16, 12), .Cells(16, 12)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(16, 12), .Cells(16, 12)).VerticalAlignment = xlCenter
         End If
                    
         g_rst_Princi.MoveNext
      Loop
      
      'BBP
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD IN ('021','022','023')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .Range(.Cells(18, 3), .Cells(18, 3)) = g_rst_Princi!CONT
            .Range(.Cells(18, 3), .Cells(18, 3)).VerticalAlignment = xlCenter
            .Range(.Cells(18, 4), .Cells(18, 4)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(18, 4), .Cells(18, 4)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .Range(.Cells(18, 5), .Cells(18, 5)) = g_rst_Princi!CONT
            .Range(.Cells(18, 5), .Cells(18, 5)).VerticalAlignment = xlCenter
            .Range(.Cells(18, 6), .Cells(18, 6)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(18, 6), .Cells(18, 6)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .Range(.Cells(18, 7), .Cells(18, 7)) = g_rst_Princi!CONT
            .Range(.Cells(18, 7), .Cells(18, 7)).VerticalAlignment = xlCenter
            .Range(.Cells(18, 8), .Cells(18, 8)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(18, 8), .Cells(18, 8)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .Range(.Cells(18, 9), .Cells(18, 9)) = g_rst_Princi!CONT
            .Range(.Cells(18, 9), .Cells(18, 9)).VerticalAlignment = xlCenter
            .Range(.Cells(18, 10), .Cells(18, 10)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(18, 10), .Cells(18, 10)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .Range(.Cells(18, 11), .Cells(18, 11)) = g_rst_Princi!CONT
            .Range(.Cells(18, 11), .Cells(18, 11)).VerticalAlignment = xlCenter
            .Range(.Cells(18, 12), .Cells(18, 12)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(18, 12), .Cells(18, 12)).VerticalAlignment = xlCenter
         End If
                    
         g_rst_Princi.MoveNext
      Loop
      
      'TECHO PROPIO
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD = '024' "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .Range(.Cells(20, 3), .Cells(20, 3)) = g_rst_Princi!CONT
            .Range(.Cells(20, 3), .Cells(20, 3)).VerticalAlignment = xlCenter
            .Range(.Cells(20, 4), .Cells(20, 4)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(20, 4), .Cells(20, 4)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .Range(.Cells(20, 5), .Cells(20, 5)) = g_rst_Princi!CONT
            .Range(.Cells(20, 5), .Cells(20, 5)).VerticalAlignment = xlCenter
            .Range(.Cells(20, 6), .Cells(20, 6)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(20, 6), .Cells(20, 6)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .Range(.Cells(20, 7), .Cells(20, 7)) = g_rst_Princi!CONT
            .Range(.Cells(20, 7), .Cells(20, 7)).VerticalAlignment = xlCenter
            .Range(.Cells(20, 8), .Cells(20, 8)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(20, 8), .Cells(20, 8)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .Range(.Cells(20, 9), .Cells(20, 9)) = g_rst_Princi!CONT
            .Range(.Cells(20, 9), .Cells(20, 9)).VerticalAlignment = xlCenter
            .Range(.Cells(20, 10), .Cells(20, 10)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(20, 10), .Cells(20, 10)).VerticalAlignment = xlCenter
         End If
         
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .Range(.Cells(20, 11), .Cells(20, 11)) = g_rst_Princi!CONT
            .Range(.Cells(20, 11), .Cells(20, 11)).VerticalAlignment = xlCenter
            .Range(.Cells(20, 12), .Cells(20, 12)) = CLng(g_rst_Princi!SALDO)
            .Range(.Cells(20, 12), .Cells(20, 12)).VerticalAlignment = xlCenter
         End If
                    
         g_rst_Princi.MoveNext
      Loop
      
      .Cells(8, 13).Formula = "=C8+E8+G8+I8+K8"
      .Cells(9, 13).Formula = "=D8+F8+H8+J8+L8"
      .Cells(10, 13).Formula = "=C10+E10+G10+I10+K10"
      .Cells(11, 13).Formula = "=D10+F10+H10+J10+L10"
      .Cells(12, 13).Formula = "=C12+E12+G12+I12+K12"
      .Cells(13, 13).Formula = "=D12+F12+H12+J12+L12"
      .Cells(14, 13).Formula = "=C14+E14+G14+I14+K14"
      .Cells(15, 13).Formula = "=D14+F14+H14+J14+L14"
      .Cells(16, 13).Formula = "=C16+E16+G16+I16+K16"
      .Cells(17, 13).Formula = "=D16+F16+H16+J16+L16"
      .Cells(18, 13).Formula = "=C18+E18+G18+I18+K18"
      .Cells(19, 13).Formula = "=D18+F18+H18+J18+L18"
      .Cells(20, 13).Formula = "=C20+E20+G20+I20+K20"
      .Cells(21, 13).Formula = "=D20+F20+H20+J20+L20"
      
      .Cells(24, 3).Formula = "=C8+C10+C12+C14+C16+C18+C20"
      .Cells(24, 4).Formula = "=D8+D10+D12+D14+D16+D18+D20"
      .Cells(24, 5).Formula = "=E8+E10+E12+E14+E16+E18+E20"
      .Cells(24, 6).Formula = "=F8+F10+F12+F14+F16+F18+F20"
      .Cells(24, 7).Formula = "=G8+G10+G12+G14+G16+G18+G20"
      .Cells(24, 8).Formula = "=H8+H10+H12+H14+H16+H18+H20"
      .Cells(24, 9).Formula = "=I8+I10+I12+I14+I16+I18+I20"
      .Cells(24, 10).Formula = "=J8+J10+J12+J14+J16+J18+J20"
      .Cells(24, 11).Formula = "=K8+K10+K12+K14+K16+K18+K20"
      .Cells(24, 12).Formula = "=L8+L10+L12+L14+L16+L18+L20"
      
      .Cells(22, 13).Formula = "=M8+M10+M12+M14+M16+M18+M20"
      .Cells(23, 13).Formula = "=M9+M11+M13+M15+M17+M19+M21"
      
      If .Cells(8, 13) > 0 Then
         .Cells(9, 3).Formula = "=C8/M8"
         .Cells(9, 4).Formula = "=D8/M9"
         .Cells(9, 5).Formula = "=E8/M8"
         .Cells(9, 6).Formula = "=F8/M9"
         .Cells(9, 7).Formula = "=G8/M8"
         .Cells(9, 8).Formula = "=H8/M9"
         .Cells(9, 9).Formula = "=I8/M8"
         .Cells(9, 10).Formula = "=J8/M9"
         .Cells(9, 11).Formula = "=K8/M8"
         .Cells(9, 12).Formula = "=L8/M9"
      End If
      If .Cells(10, 13) > 0 Then
         .Cells(11, 3).Formula = "=C10/M10"
         .Cells(11, 4).Formula = "=D10/M11"
         .Cells(11, 5).Formula = "=E10/M10"
         .Cells(11, 6).Formula = "=F10/M11"
         .Cells(11, 7).Formula = "=G10/M10"
         .Cells(11, 8).Formula = "=H10/M11"
         .Cells(11, 9).Formula = "=I10/M10"
         .Cells(11, 10).Formula = "=J10/M11"
         .Cells(11, 11).Formula = "=K10/M10"
         .Cells(11, 12).Formula = "=L10/M11"
      End If
      If .Cells(12, 13) > 0 Then
         .Cells(13, 3).Formula = "=C12/M12"
         .Cells(13, 4).Formula = "=D12/M13"
         .Cells(13, 5).Formula = "=E12/M12"
         .Cells(13, 6).Formula = "=F12/M13"
         .Cells(13, 7).Formula = "=G12/M12"
         .Cells(13, 8).Formula = "=H12/M13"
         .Cells(13, 9).Formula = "=I12/M12"
         .Cells(13, 10).Formula = "=J12/M13"
         .Cells(13, 11).Formula = "=K12/M12"
         .Cells(13, 12).Formula = "=L12/M13"
      End If
      If .Cells(14, 13) > 0 Then
         .Cells(15, 3).Formula = "=C14/M14"
         .Cells(15, 4).Formula = "=D14/M15"
         .Cells(15, 5).Formula = "=E14/M14"
         .Cells(15, 6).Formula = "=F14/M15"
         .Cells(15, 7).Formula = "=G14/M14"
         .Cells(15, 8).Formula = "=H14/M15"
         .Cells(15, 9).Formula = "=I14/M14"
         .Cells(15, 10).Formula = "=J14/M15"
         .Cells(15, 11).Formula = "=K14/M14"
         .Cells(15, 12).Formula = "=L14/M15"
      End If
      If .Cells(16, 13) > 0 Then
         .Cells(17, 3).Formula = "=C16/M16"
         .Cells(17, 4).Formula = "=D16/M17"
         .Cells(17, 5).Formula = "=E16/M16"
         .Cells(17, 6).Formula = "=F16/M17"
         .Cells(17, 7).Formula = "=G16/M16"
         .Cells(17, 8).Formula = "=H16/M17"
         .Cells(17, 9).Formula = "=I16/M16"
         .Cells(17, 10).Formula = "=J16/M17"
         .Cells(17, 11).Formula = "=K16/M16"
         .Cells(17, 12).Formula = "=L16/M17"
      End If
      If .Cells(18, 13) > 0 Then
         .Cells(19, 3).Formula = "=C18/M18"
         .Cells(19, 4).Formula = "=D18/M19"
         .Cells(19, 5).Formula = "=E18/M18"
         .Cells(19, 6).Formula = "=F18/M19"
         .Cells(19, 7).Formula = "=G18/M18"
         .Cells(19, 8).Formula = "=H18/M19"
         .Cells(19, 9).Formula = "=I18/M18"
         .Cells(19, 10).Formula = "=J18/M19"
         .Cells(19, 11).Formula = "=K18/M18"
         .Cells(19, 12).Formula = "=L18/M19"
      End If
      If .Cells(20, 13) > 0 Then
         .Cells(21, 3).Formula = "=C20/M20"
         .Cells(21, 4).Formula = "=D20/M21"
         .Cells(21, 5).Formula = "=E20/M20"
         .Cells(21, 6).Formula = "=F20/M21"
         .Cells(21, 7).Formula = "=G20/M20"
         .Cells(21, 8).Formula = "=H20/M21"
         .Cells(21, 9).Formula = "=I20/M20"
         .Cells(21, 10).Formula = "=J20/M21"
         .Cells(21, 11).Formula = "=K20/M20"
         .Cells(21, 12).Formula = "=L20/M21"
      End If
      
      .Range(.Cells(22, 3), .Cells(23, 3)).Merge
      .Range(.Cells(22, 4), .Cells(23, 4)).Merge
      .Range(.Cells(22, 5), .Cells(23, 5)).Merge
      .Range(.Cells(22, 6), .Cells(23, 6)).Merge
      .Range(.Cells(22, 7), .Cells(23, 7)).Merge
      .Range(.Cells(22, 8), .Cells(23, 8)).Merge
      .Range(.Cells(22, 9), .Cells(23, 9)).Merge
      .Range(.Cells(22, 10), .Cells(23, 10)).Merge
      .Range(.Cells(22, 11), .Cells(23, 11)).Merge
      .Range(.Cells(22, 12), .Cells(23, 12)).Merge
      .Range(.Cells(22, 3), .Cells(23, 12)).VerticalAlignment = xlCenter
      
      If .Cells(22, 13) > 0 Then
         .Cells(22, 3).Formula = "=C24/M22"
         .Cells(22, 4).Formula = "=D24/M23"
         .Cells(22, 5).Formula = "=E24/M22"
         .Cells(22, 6).Formula = "=F24/M23"
         .Cells(22, 7).Formula = "=G24/M22"
         .Cells(22, 8).Formula = "=H24/M23"
         .Cells(22, 9).Formula = "=I24/M22"
         .Cells(22, 10).Formula = "=J24/M23"
         .Cells(22, 11).Formula = "=K24/M22"
         .Cells(22, 12).Formula = "=L24/M23"
      End If
      
      .Range(.Cells(22, 2), .Cells(23, 13)).Font.Bold = True
      .Range(.Cells(24, 2), .Cells(24, 13)).Font.Bold = True
      
      .Range(.Cells(9, 3), .Cells(9, 12)).NumberFormat = "0.00%"
      .Range(.Cells(11, 3), .Cells(11, 12)).NumberFormat = "0.00%"
      .Range(.Cells(13, 3), .Cells(13, 12)).NumberFormat = "0.00%"
      .Range(.Cells(15, 3), .Cells(15, 12)).NumberFormat = "0.00%"
      .Range(.Cells(17, 3), .Cells(17, 12)).NumberFormat = "0.00%"
      .Range(.Cells(19, 3), .Cells(19, 12)).NumberFormat = "0.00%"
      .Range(.Cells(21, 3), .Cells(21, 12)).NumberFormat = "0.00%"
      .Range(.Cells(22, 3), .Cells(22, 12)).NumberFormat = "0.00%"
            
      For r_int_Contad = 8 To 24
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 3), .Cells(r_int_Contad, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
   End With
End Sub

Private Sub fs_GenExc_Detalle_Clasificacion()
   Dim r_int_Contad     As Integer
   Dim r_int_fila       As Integer
   Dim r_Cad_Ante       As String
  
   With r_obj_Excel.Sheets(2)
      .Name = "DET.CLASIF."
      'Titulo
      .Cells(2, 2) = "DETALLE POR CLASIFICACION DEL MES " & UCase(Trim(cmb_PerMes.Text)) & " DEL " & CStr(ipp_PerAno.Text)
      .Range(.Cells(2, 2), .Cells(2, 7)).Merge
      .Range("B2:G2").HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 2), .Cells(2, 7)).Font.Name = "Calibri"
      .Range(.Cells(2, 2), .Cells(2, 7)).Font.Size = 12
      .Range(.Cells(2, 2), .Cells(2, 7)).Font.Bold = True
               
      .Columns("D").NumberFormat = "###,##0"
      .Columns("E").NumberFormat = "###,##0"
      .Columns("F").NumberFormat = "###,##0.00"
      
      .Columns("A").ColumnWidth = 4
      .Columns("B").ColumnWidth = 20
      .Columns("C").ColumnWidth = 20
      .Columns("D").ColumnWidth = 12
      .Columns("E").ColumnWidth = 10
      .Columns("F").ColumnWidth = 15
      .Columns("G").ColumnWidth = 10
      
      .Range(.Cells(4, 2), .Cells(4, 2)) = "Estado"
      .Range(.Cells(4, 3), .Cells(4, 3)) = "Situacion"
      .Range(.Cells(4, 4), .Cells(4, 4)) = "N° Creditos"
      .Range(.Cells(4, 5), .Cells(4, 5)) = "%"
      .Range(.Cells(4, 6), .Cells(4, 6)) = "Saldo"
      .Range(.Cells(4, 7), .Cells(4, 7)) = "%"
            
      .Range(.Cells(4, 2), .Cells(4, 7)).HorizontalAlignment = xlVAlignCenter
      .Range(.Cells(4, 2), .Cells(4, 7)).Font.Name = "Calibri"
      .Range(.Cells(4, 2), .Cells(4, 7)).Font.Size = 10
      .Range(.Cells(4, 2), .Cells(4, 7)).Font.Bold = True
      
      .Range(.Cells(4, 2), .Cells(4, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(4, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(4, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(4, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(4, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(5, 2), .Cells(8, 2)).Merge
      .Range(.Cells(5, 2), .Cells(8, 2)) = "DEFICIENTE"
      .Range(.Cells(5, 2), .Cells(8, 2)).VerticalAlignment = xlCenter
      
      .Range(.Cells(9, 2), .Cells(12, 2)).Merge
      .Range(.Cells(9, 2), .Cells(12, 2)) = "DUDOSO"
      .Range(.Cells(9, 2), .Cells(12, 2)).VerticalAlignment = xlCenter
      
      .Range(.Cells(13, 2), .Cells(16, 2)).Merge
      .Range(.Cells(13, 2), .Cells(16, 2)) = "PERDIDA"
      .Range(.Cells(13, 2), .Cells(16, 2)).VerticalAlignment = xlCenter

      .Range(.Cells(5, 3), .Cells(5, 3)) = "Alineado"
      .Range(.Cells(6, 3), .Cells(6, 3)) = "Alineado Moroso"
      .Range(.Cells(7, 3), .Cells(7, 3)) = "Moroso"
      .Range(.Cells(8, 3), .Cells(8, 3)) = "Total"
      
      .Range(.Cells(9, 3), .Cells(9, 3)) = "Alineado"
      .Range(.Cells(10, 3), .Cells(10, 3)) = "Alineado Moroso"
      .Range(.Cells(11, 3), .Cells(11, 3)) = "Moroso"
      .Range(.Cells(12, 3), .Cells(12, 3)) = "Total"
      
      .Range(.Cells(13, 3), .Cells(13, 3)) = "Alineado"
      .Range(.Cells(14, 3), .Cells(14, 3)) = "Alineado Moroso"
      .Range(.Cells(15, 3), .Cells(15, 3)) = "Moroso"
      .Range(.Cells(16, 3), .Cells(16, 3)) = "Total"

      'Procesa infomacion
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "USP_RPT_CLASIFICACION("
      g_str_Parame = g_str_Parame & CInt(r_int_PerAno) & ", "
      g_str_Parame = g_str_Parame & CInt(r_int_PerMes) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'REPORTE_CLASIFICACION', "
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         grd_LisDetCalif.Redraw = True
         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
' r_obj_Excel.Visible = True
      r_int_fila = 5
      
      Do While Not g_rst_Princi.EOF
         r_Cad_Ante = g_rst_Princi!RPT_VALCAD01
         If g_rst_Princi!RPT_VALCAD01 = 2 Then
            .Range(.Cells(r_int_fila, 4), .Cells(r_int_fila, 4)) = Format(g_rst_Princi!RPT_VALNUM01, "#,##0")
            .Range(.Cells(r_int_fila, 5), .Cells(r_int_fila, 5)) = Format(g_rst_Princi!RPT_VALNUM02, "##0.00")
            .Range(.Cells(r_int_fila, 6), .Cells(r_int_fila, 6)) = Format(g_rst_Princi!RPT_VALNUM03, "###,##0.00")
            .Range(.Cells(r_int_fila, 7), .Cells(r_int_fila, 7)) = Format(g_rst_Princi!RPT_VALNUM04, "##0.00")
         End If
         If g_rst_Princi!RPT_VALCAD01 = 3 Then
            .Range(.Cells(r_int_fila, 4), .Cells(r_int_fila, 4)) = Format(g_rst_Princi!RPT_VALNUM01, "#,##0")
            .Range(.Cells(r_int_fila, 5), .Cells(r_int_fila, 5)) = Format(g_rst_Princi!RPT_VALNUM02, "##0.00")
            .Range(.Cells(r_int_fila, 6), .Cells(r_int_fila, 6)) = Format(g_rst_Princi!RPT_VALNUM03, "###,##0.00")
            .Range(.Cells(r_int_fila, 7), .Cells(r_int_fila, 7)) = Format(g_rst_Princi!RPT_VALNUM04, "##0.00")
         End If
         If g_rst_Princi!RPT_VALCAD01 = 4 Then
            .Range(.Cells(r_int_fila, 4), .Cells(r_int_fila, 4)) = Format(g_rst_Princi!RPT_VALNUM01, "#,##0")
            .Range(.Cells(r_int_fila, 5), .Cells(r_int_fila, 5)) = Format(g_rst_Princi!RPT_VALNUM02, "##0.00")
            .Range(.Cells(r_int_fila, 6), .Cells(r_int_fila, 6)) = Format(g_rst_Princi!RPT_VALNUM03, "###,##0.00")
            .Range(.Cells(r_int_fila, 7), .Cells(r_int_fila, 7)) = Format(g_rst_Princi!RPT_VALNUM04, "##0.00")
         End If
                 
         r_int_fila = r_int_fila + 1
         g_rst_Princi.MoveNext
         
         If Not g_rst_Princi.EOF Then
            If r_Cad_Ante <> g_rst_Princi!RPT_VALCAD01 Then
               r_int_fila = r_int_fila + 1
            End If
         End If
      Loop

      'CREDITOS
      .Cells(8, 4).Formula = "=D5+D6+D7"
      .Cells(12, 4).Formula = "=D9+D10+D11"
      .Cells(16, 4).Formula = "=D13+D14+D15"
      
      If .Cells(8, 4) > 0 Then
         .Cells(5, 5) = grd_LisDetCalif.TextMatrix(1, 3)
         .Cells(6, 5) = grd_LisDetCalif.TextMatrix(2, 3)
         .Cells(7, 5) = grd_LisDetCalif.TextMatrix(3, 3)
      End If
      If .Cells(12, 4) > 0 Then
         .Cells(9, 5) = grd_LisDetCalif.TextMatrix(5, 3)
         .Cells(10, 5) = grd_LisDetCalif.TextMatrix(6, 3)
         .Cells(11, 5) = grd_LisDetCalif.TextMatrix(7, 3)
      End If
      If .Cells(16, 4) > 0 Then
         .Cells(13, 5) = grd_LisDetCalif.TextMatrix(9, 3)
         .Cells(14, 5) = grd_LisDetCalif.TextMatrix(10, 3)
         .Cells(15, 5) = grd_LisDetCalif.TextMatrix(11, 3)
      End If
       
      .Cells(8, 5).Formula = "=E5+E6+E7"
      .Cells(12, 5).Formula = "=E9+E10+E11"
      .Cells(16, 5).Formula = "=E13+E14+E15"
      
      .Range(.Cells(5, 5), .Cells(16, 5)).NumberFormat = "0.00%"
      
      'SALDO
      .Cells(8, 6).Formula = "=F5+F6+F7"
      .Cells(12, 6).Formula = "=F9+F10+F11"
      .Cells(16, 6).Formula = "=F13+F14+F15"
      
      If .Cells(8, 6) > 0 Then
         .Cells(5, 7) = grd_LisDetCalif.TextMatrix(1, 5)
         .Cells(6, 7) = grd_LisDetCalif.TextMatrix(2, 5)
         .Cells(7, 7) = grd_LisDetCalif.TextMatrix(3, 5)
      End If
      If .Cells(12, 6) > 0 Then
         .Cells(9, 7) = grd_LisDetCalif.TextMatrix(5, 5)
         .Cells(10, 7) = grd_LisDetCalif.TextMatrix(6, 5)
         .Cells(11, 7) = grd_LisDetCalif.TextMatrix(7, 5)
      End If
      If .Cells(16, 6) > 0 Then
         .Cells(13, 7) = grd_LisDetCalif.TextMatrix(9, 5)
         .Cells(14, 7) = grd_LisDetCalif.TextMatrix(10, 5)
         .Cells(15, 7) = grd_LisDetCalif.TextMatrix(11, 5)
      End If
      
      .Cells(8, 7).Formula = "=G5+G6+G7"
      .Cells(12, 7).Formula = "=G9+G10+G11"
      .Cells(16, 7).Formula = "=G13+G14+G15"
      
      .Range(.Cells(5, 7), .Cells(16, 7)).NumberFormat = "0.00%"
      .Range(.Cells(8, 3), .Cells(8, 7)).Font.Bold = True
      .Range(.Cells(12, 3), .Cells(12, 7)).Font.Bold = True
      .Range(.Cells(16, 3), .Cells(16, 7)).Font.Bold = True
      .Range(.Cells(8, 3), .Cells(8, 7)).Interior.Color = RGB(238, 238, 238)
      .Range(.Cells(12, 3), .Cells(12, 7)).Interior.Color = RGB(238, 238, 238)
      .Range(.Cells(16, 3), .Cells(16, 7)).Interior.Color = RGB(238, 238, 238)
      
      For r_int_Contad = 5 To 16
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 3), .Cells(r_int_Contad, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
   End With
End Sub

Private Sub fs_GenExc_Detalle_Producto()
   Dim r_int_Contad     As Integer
   Dim r_int_Cant       As Integer
   Dim r_dbl_Monto      As Double
   
   With r_obj_Excel.Sheets(3)
      'Titulo
      .Name = "DET.PROD."
      .Cells(2, 2) = "DETALLE POR PRODUCTO DEL MES " & UCase(Trim(cmb_PerMes.Text)) & " DEL " & CStr(ipp_PerAno.Text)
      .Range(.Cells(2, 2), .Cells(2, 11)).Merge
      .Range("B2:M2").HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 2), .Cells(2, 11)).Font.Name = "Calibri"
      .Range(.Cells(2, 2), .Cells(2, 11)).Font.Size = 12
      .Range(.Cells(2, 2), .Cells(2, 11)).Font.Bold = True
               
      .Columns("D").NumberFormat = "###,##0"
      .Columns("E").NumberFormat = "###,##0.00"
      .Columns("G").NumberFormat = "###,##0"
      .Columns("H").NumberFormat = "###,##0.00"
      .Columns("J").NumberFormat = "###,##0"
      .Columns("K").NumberFormat = "###,##0.00"
      
      .Columns("A").ColumnWidth = 4
      .Columns("B").ColumnWidth = 20
      .Columns("C").ColumnWidth = 20
      .Columns("D").ColumnWidth = 10
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 20
      .Columns("G").ColumnWidth = 10
      .Columns("H").ColumnWidth = 15
      .Columns("I").ColumnWidth = 20
      .Columns("J").ColumnWidth = 10
      .Columns("K").ColumnWidth = 15
      
      .Range(.Cells(4, 2), .Cells(5, 2)).Merge
      .Range(.Cells(4, 2), .Cells(5, 2)) = "Producto"
      .Range(.Cells(4, 2), .Cells(5, 2)).WrapText = True
      .Range(.Cells(4, 2), .Cells(5, 2)).VerticalAlignment = xlCenter
      
      .Range(.Cells(4, 3), .Cells(4, 5)) = "Deficiente"
      .Range(.Cells(4, 3), .Cells(4, 5)).Merge
      .Range(.Cells(4, 6), .Cells(4, 8)) = "Dudoso"
      .Range(.Cells(4, 6), .Cells(4, 8)).Merge
      .Range(.Cells(4, 9), .Cells(4, 11)) = "Perdida"
      .Range(.Cells(4, 9), .Cells(4, 11)).Merge
      
      .Range(.Cells(5, 3), .Cells(5, 3)) = "Situacion"
      .Range(.Cells(5, 4), .Cells(5, 4)) = "N° Creditos"
      .Range(.Cells(5, 5), .Cells(5, 5)) = "Saldo"
      
      .Range(.Cells(5, 6), .Cells(5, 6)) = "Situacion"
      .Range(.Cells(5, 7), .Cells(5, 7)) = "N° Creditos"
      .Range(.Cells(5, 8), .Cells(5, 8)) = "Saldo"
      
      .Range(.Cells(5, 9), .Cells(5, 9)) = "Situacion"
      .Range(.Cells(5, 10), .Cells(5, 10)) = "N° Creditos"
      .Range(.Cells(5, 11), .Cells(5, 11)) = "Saldo"
     
      .Range(.Cells(4, 2), .Cells(5, 11)).HorizontalAlignment = xlVAlignCenter
      .Range(.Cells(4, 2), .Cells(5, 11)).Font.Name = "Calibri"
      .Range(.Cells(4, 2), .Cells(5, 11)).Font.Size = 10
      .Range(.Cells(4, 2), .Cells(5, 11)).Font.Bold = True
      
      .Range(.Cells(4, 2), .Cells(5, 11)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(5, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(5, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(6, 3), .Cells(5, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(5, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(5, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(5, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(6, 2), .Cells(9, 2)).Merge
      .Range(.Cells(6, 2), .Cells(9, 2)) = "CRC-PBP"
      .Range(.Cells(6, 2), .Cells(9, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(10, 2), .Cells(13, 2)).Merge
      .Range(.Cells(10, 2), .Cells(13, 2)) = "MICASITA"
      .Range(.Cells(10, 2), .Cells(13, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(14, 2), .Cells(17, 2)).Merge
      .Range(.Cells(14, 2), .Cells(17, 2)) = "CME"
      .Range(.Cells(14, 2), .Cells(17, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(18, 2), .Cells(21, 2)).Merge
      .Range(.Cells(18, 2), .Cells(21, 2)) = "N. MIVIVIENDA"
      .Range(.Cells(18, 2), .Cells(21, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(22, 2), .Cells(25, 2)).Merge
      .Range(.Cells(22, 2), .Cells(25, 2)) = "MICASA MAS"
      .Range(.Cells(22, 2), .Cells(25, 2)).VerticalAlignment = xlCenter
      .Range(.Cells(26, 2), .Cells(29, 2)).Merge
      .Range(.Cells(26, 2), .Cells(29, 2)) = "BBP"
      .Range(.Cells(26, 2), .Cells(29, 2)).VerticalAlignment = xlCenter
      
      .Range(.Cells(30, 2), .Cells(33, 2)).Merge
      .Range(.Cells(30, 2), .Cells(33, 2)) = "TECHO PROPIO"
      .Range(.Cells(30, 2), .Cells(33, 2)).VerticalAlignment = xlCenter
      
      For r_int_Contad = 3 To 9
         .Range(.Cells(6, r_int_Contad), .Cells(6, r_int_Contad)) = "ALINEADO"
         .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)) = "ALINEADO MOROSO"
         .Range(.Cells(8, r_int_Contad), .Cells(8, r_int_Contad)) = "MOROSO"
         .Range(.Cells(9, r_int_Contad), .Cells(9, r_int_Contad)) = "SUBTOTAL 1"
         r_int_Contad = r_int_Contad + 2
      Next
      For r_int_Contad = 3 To 9
         .Range(.Cells(10, r_int_Contad), .Cells(10, r_int_Contad)) = "ALINEADO"
         .Range(.Cells(11, r_int_Contad), .Cells(11, r_int_Contad)) = "ALINEADO MOROSO"
         .Range(.Cells(12, r_int_Contad), .Cells(12, r_int_Contad)) = "MOROSO"
         .Range(.Cells(13, r_int_Contad), .Cells(13, r_int_Contad)) = "SUBTOTAL 2"
         r_int_Contad = r_int_Contad + 2
      Next
      For r_int_Contad = 3 To 9
         .Range(.Cells(14, r_int_Contad), .Cells(14, r_int_Contad)) = "ALINEADO"
         .Range(.Cells(15, r_int_Contad), .Cells(15, r_int_Contad)) = "ALINEADO MOROSO"
         .Range(.Cells(16, r_int_Contad), .Cells(16, r_int_Contad)) = "MOROSO"
         .Range(.Cells(17, r_int_Contad), .Cells(17, r_int_Contad)) = "SUBTOTAL 3"
         r_int_Contad = r_int_Contad + 2
      Next
      For r_int_Contad = 3 To 9
         .Range(.Cells(18, r_int_Contad), .Cells(18, r_int_Contad)) = "ALINEADO"
         .Range(.Cells(19, r_int_Contad), .Cells(19, r_int_Contad)) = "ALINEADO MOROSO"
         .Range(.Cells(20, r_int_Contad), .Cells(20, r_int_Contad)) = "MOROSO"
         .Range(.Cells(21, r_int_Contad), .Cells(21, r_int_Contad)) = "SUBTOTAL 4"
         r_int_Contad = r_int_Contad + 2
      Next
      For r_int_Contad = 3 To 9
         .Range(.Cells(22, r_int_Contad), .Cells(22, r_int_Contad)) = "ALINEADO"
         .Range(.Cells(23, r_int_Contad), .Cells(23, r_int_Contad)) = "ALINEADO MOROSO"
         .Range(.Cells(24, r_int_Contad), .Cells(24, r_int_Contad)) = "MOROSO"
         .Range(.Cells(25, r_int_Contad), .Cells(25, r_int_Contad)) = "SUBTOTAL 5"
         r_int_Contad = r_int_Contad + 2
      Next
      For r_int_Contad = 3 To 9
         .Range(.Cells(26, r_int_Contad), .Cells(26, r_int_Contad)) = "ALINEADO"
         .Range(.Cells(27, r_int_Contad), .Cells(27, r_int_Contad)) = "ALINEADO MOROSO"
         .Range(.Cells(28, r_int_Contad), .Cells(28, r_int_Contad)) = "MOROSO"
         .Range(.Cells(29, r_int_Contad), .Cells(29, r_int_Contad)) = "SUBTOTAL 6"
         r_int_Contad = r_int_Contad + 2
      Next
      For r_int_Contad = 3 To 9
         .Range(.Cells(30, r_int_Contad), .Cells(30, r_int_Contad)) = "ALINEADO"
         .Range(.Cells(31, r_int_Contad), .Cells(31, r_int_Contad)) = "ALINEADO MOROSO"
         .Range(.Cells(32, r_int_Contad), .Cells(32, r_int_Contad)) = "MOROSO"
         .Range(.Cells(33, r_int_Contad), .Cells(33, r_int_Contad)) = "SUBTOTAL 7"
         r_int_Contad = r_int_Contad + 2
      Next
      For r_int_Contad = 3 To 9
         .Range(.Cells(34, r_int_Contad), .Cells(34, r_int_Contad)) = "TOTAL"
         r_int_Contad = r_int_Contad + 2
      Next
      
      .Range(.Cells(9, 3), .Cells(9, 11)).Font.Bold = True
      .Range(.Cells(13, 3), .Cells(13, 11)).Font.Bold = True
      .Range(.Cells(17, 3), .Cells(17, 11)).Font.Bold = True
      .Range(.Cells(21, 3), .Cells(21, 11)).Font.Bold = True
      .Range(.Cells(25, 3), .Cells(25, 11)).Font.Bold = True
      .Range(.Cells(29, 3), .Cells(29, 11)).Font.Bold = True
      .Range(.Cells(33, 3), .Cells(33, 11)).Font.Bold = True
      .Range(.Cells(34, 3), .Cells(34, 11)).Font.Bold = True
      
      .Range(.Cells(9, 3), .Cells(9, 11)).Interior.Color = RGB(238, 238, 238)
      .Range(.Cells(13, 3), .Cells(13, 11)).Interior.Color = RGB(238, 238, 238)
      .Range(.Cells(17, 3), .Cells(17, 11)).Interior.Color = RGB(238, 238, 238)
      .Range(.Cells(21, 3), .Cells(21, 11)).Interior.Color = RGB(238, 238, 238)
      .Range(.Cells(25, 3), .Cells(25, 11)).Interior.Color = RGB(238, 238, 238)
      .Range(.Cells(29, 3), .Cells(29, 11)).Interior.Color = RGB(238, 238, 238)
      .Range(.Cells(33, 2), .Cells(33, 11)).Interior.Color = RGB(238, 238, 238)
      .Range(.Cells(34, 2), .Cells(34, 11)).Interior.Color = RGB(238, 238, 238)

      '----------------------------------------------------------------------------
      'DEFICIENTE(CRC-PBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrCRC & ") "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(6, 4), .Cells(6, 4)) = g_rst_Princi!CONT
            .Range(.Cells(6, 5), .Cells(6, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .Range(.Cells(7, 4), .Cells(7, 4)) = g_rst_Princi!CONT
            .Range(.Cells(7, 5), .Cells(7, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .Range(.Cells(8, 4), .Cells(8, 4)) = g_rst_Princi!CONT
            .Range(.Cells(8, 5), .Cells(8, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DEFICIENTE(MICASITA)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(10, 4), .Cells(10, 4)) = g_rst_Princi!CONT
            .Range(.Cells(10, 5), .Cells(10, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .Range(.Cells(11, 4), .Cells(11, 4)) = g_rst_Princi!CONT
            .Range(.Cells(11, 5), .Cells(11, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .Range(.Cells(12, 4), .Cells(12, 4)) = g_rst_Princi!CONT
            .Range(.Cells(12, 5), .Cells(12, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'DEFICIENTE(CME)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrCME & ") "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(14, 4), .Cells(14, 4)) = g_rst_Princi!CONT
            .Range(.Cells(14, 5), .Cells(14, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .Range(.Cells(15, 4), .Cells(15, 4)) = g_rst_Princi!CONT
            .Range(.Cells(15, 5), .Cells(15, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .Range(.Cells(16, 4), .Cells(16, 4)) = g_rst_Princi!CONT
            .Range(.Cells(16, 5), .Cells(16, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DEFICIENTE(N.MIVIVIENDA)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND ( HIPCIE_CODPRD IN (" & moddat_g_str_AgrMIHG & "," & moddat_g_str_Agr2FMV & ") OR HIPCIE_CODPRD = '025') "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(18, 4), .Cells(18, 4)) = g_rst_Princi!CONT
            .Range(.Cells(18, 5), .Cells(18, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .Range(.Cells(19, 4), .Cells(19, 4)) = g_rst_Princi!CONT
            .Range(.Cells(19, 5), .Cells(19, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .Range(.Cells(20, 4), .Cells(20, 4)) = g_rst_Princi!CONT
            .Range(.Cells(20, 5), .Cells(20, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'DEFICIENTE(MICASAMAS)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD = '019'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(22, 4), .Cells(22, 4)) = g_rst_Princi!CONT
            .Range(.Cells(22, 5), .Cells(22, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .Range(.Cells(23, 4), .Cells(23, 4)) = g_rst_Princi!CONT
            .Range(.Cells(23, 5), .Cells(23, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .Range(.Cells(24, 4), .Cells(24, 4)) = g_rst_Princi!CONT
            .Range(.Cells(24, 5), .Cells(24, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DEFICIENTE(BBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD IN ('021','022','023')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(26, 4), .Cells(26, 4)) = g_rst_Princi!CONT
            .Range(.Cells(26, 5), .Cells(26, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .Range(.Cells(27, 4), .Cells(27, 4)) = g_rst_Princi!CONT
            .Range(.Cells(27, 5), .Cells(27, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .Range(.Cells(28, 4), .Cells(28, 4)) = g_rst_Princi!CONT
            .Range(.Cells(28, 5), .Cells(28, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'DEFICIENTE(TECHO PROPIO)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD = '024' "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(30, 4), .Cells(30, 4)) = g_rst_Princi!CONT
            .Range(.Cells(30, 5), .Cells(30, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .Range(.Cells(31, 4), .Cells(31, 4)) = g_rst_Princi!CONT
            .Range(.Cells(31, 5), .Cells(31, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .Range(.Cells(32, 4), .Cells(32, 4)) = g_rst_Princi!CONT
            .Range(.Cells(32, 5), .Cells(32, 5)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      '----------------------------------------------------------------------------
      'DUDOSO(CRC-PBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrCRC & ") "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      'r_obj_Excel.Visible = True
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(6, 7), .Cells(6, 7)) = .Range(.Cells(6, 7), .Cells(6, 7)) + g_rst_Princi!CONT
            .Range(.Cells(6, 8), .Cells(6, 8)) = .Range(.Cells(6, 8), .Cells(6, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(7, 7), .Cells(7, 7)) = .Range(.Cells(7, 7), .Cells(7, 7)) + g_rst_Princi!CONT
            .Range(.Cells(7, 8), .Cells(7, 8)) = .Range(.Cells(7, 8), .Cells(7, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .Range(.Cells(8, 7), .Cells(8, 7)) = .Range(.Cells(8, 7), .Cells(8, 7)) + g_rst_Princi!CONT
            .Range(.Cells(8, 8), .Cells(8, 8)) = .Range(.Cells(8, 8), .Cells(8, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DUDOSO(MICASITA)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(10, 7), .Cells(10, 7)) = .Range(.Cells(10, 7), .Cells(10, 7)) + g_rst_Princi!CONT
            .Range(.Cells(10, 8), .Cells(10, 8)) = .Range(.Cells(10, 8), .Cells(10, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(11, 7), .Cells(11, 7)) = .Range(.Cells(11, 7), .Cells(11, 7)) + g_rst_Princi!CONT
            .Range(.Cells(11, 8), .Cells(11, 8)) = .Range(.Cells(11, 8), .Cells(11, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .Range(.Cells(12, 7), .Cells(12, 7)) = .Range(.Cells(12, 7), .Cells(12, 7)) + g_rst_Princi!CONT
            .Range(.Cells(12, 8), .Cells(12, 8)) = .Range(.Cells(12, 8), .Cells(12, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DUDOSO(CME)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrCME & ") "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(14, 7), .Cells(14, 7)) = .Range(.Cells(14, 7), .Cells(14, 7)) + g_rst_Princi!CONT
            .Range(.Cells(14, 8), .Cells(14, 8)) = .Range(.Cells(14, 8), .Cells(14, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(15, 7), .Cells(15, 7)) = .Range(.Cells(15, 7), .Cells(15, 7)) + g_rst_Princi!CONT
            .Range(.Cells(15, 8), .Cells(15, 8)) = .Range(.Cells(15, 8), .Cells(15, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .Range(.Cells(16, 7), .Cells(16, 7)) = .Range(.Cells(16, 7), .Cells(16, 7)) + g_rst_Princi!CONT
            .Range(.Cells(16, 8), .Cells(16, 8)) = .Range(.Cells(16, 8), .Cells(16, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DUDOSO(N.MIVIVIENDA)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND ( HIPCIE_CODPRD IN (" & moddat_g_str_AgrMIHG & "," & moddat_g_str_Agr2FMV & ") OR HIPCIE_CODPRD = '025' ) "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(18, 7), .Cells(18, 7)) = .Range(.Cells(18, 7), .Cells(18, 7)) + g_rst_Princi!CONT
            .Range(.Cells(18, 8), .Cells(18, 8)) = .Range(.Cells(18, 8), .Cells(18, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(19, 7), .Cells(19, 7)) = .Range(.Cells(19, 7), .Cells(19, 7)) + g_rst_Princi!CONT
            .Range(.Cells(19, 8), .Cells(19, 8)) = .Range(.Cells(19, 8), .Cells(19, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .Range(.Cells(20, 7), .Cells(20, 7)) = .Range(.Cells(20, 7), .Cells(20, 7)) + g_rst_Princi!CONT
            .Range(.Cells(20, 8), .Cells(20, 8)) = .Range(.Cells(20, 8), .Cells(20, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'DUDOSO(MICASAMAS)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD = '019'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(22, 7), .Cells(22, 7)) = .Range(.Cells(22, 7), .Cells(22, 7)) + g_rst_Princi!CONT
            .Range(.Cells(22, 8), .Cells(22, 8)) = .Range(.Cells(22, 8), .Cells(22, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(23, 7), .Cells(23, 7)) = .Range(.Cells(23, 7), .Cells(23, 7)) + g_rst_Princi!CONT
            .Range(.Cells(23, 8), .Cells(23, 8)) = .Range(.Cells(23, 8), .Cells(23, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .Range(.Cells(24, 7), .Cells(24, 7)) = .Range(.Cells(24, 7), .Cells(24, 7)) + g_rst_Princi!CONT
            .Range(.Cells(24, 8), .Cells(24, 8)) = .Range(.Cells(24, 8), .Cells(24, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DUDOSO(BBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD IN ('021','022','023')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(26, 7), .Cells(26, 7)) = .Range(.Cells(26, 7), .Cells(26, 7)) + g_rst_Princi!CONT
            .Range(.Cells(26, 8), .Cells(26, 8)) = .Range(.Cells(26, 8), .Cells(26, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(27, 7), .Cells(27, 7)) = .Range(.Cells(27, 7), .Cells(27, 7)) + g_rst_Princi!CONT
            .Range(.Cells(27, 8), .Cells(27, 8)) = .Range(.Cells(27, 8), .Cells(27, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .Range(.Cells(28, 7), .Cells(28, 7)) = .Range(.Cells(28, 7), .Cells(28, 7)) + g_rst_Princi!CONT
            .Range(.Cells(28, 8), .Cells(28, 8)) = .Range(.Cells(28, 8), .Cells(28, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'DUDOSO(TECHO PROPIO)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD = '024' "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(30, 7), .Cells(30, 7)) = .Range(.Cells(30, 7), .Cells(30, 7)) + g_rst_Princi!CONT
            .Range(.Cells(30, 8), .Cells(30, 8)) = .Range(.Cells(30, 8), .Cells(30, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(31, 7), .Cells(31, 7)) = .Range(.Cells(31, 7), .Cells(31, 7)) + g_rst_Princi!CONT
            .Range(.Cells(31, 8), .Cells(31, 8)) = .Range(.Cells(31, 8), .Cells(31, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .Range(.Cells(32, 7), .Cells(32, 7)) = .Range(.Cells(32, 7), .Cells(32, 7)) + g_rst_Princi!CONT
            .Range(.Cells(32, 8), .Cells(32, 8)) = .Range(.Cells(32, 8), .Cells(32, 8)) + Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      '----------------------------------------------------------------------------
      'PERDIDA(CRC-PBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD = '001'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(6, 10), .Cells(6, 10)) = g_rst_Princi!CONT
            .Range(.Cells(6, 11), .Cells(6, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            .Range(.Cells(7, 10), .Cells(7, 10)) = r_int_Cant
            .Range(.Cells(7, 11), .Cells(7, 11)) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(8, 10), .Cells(8, 10)) = g_rst_Princi!CONT
            .Range(.Cells(8, 11), .Cells(8, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'PERDIDA(MICASITA)
      r_int_Cant = 0
      r_dbl_Monto = 0
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD IN ('002','006','011')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(10, 10), .Cells(10, 10)) = g_rst_Princi!CONT
            .Range(.Cells(10, 11), .Cells(10, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            .Range(.Cells(11, 10), .Cells(11, 10)) = r_int_Cant
            .Range(.Cells(11, 11), .Cells(11, 11)) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(12, 10), .Cells(12, 10)) = g_rst_Princi!CONT
            .Range(.Cells(12, 11), .Cells(12, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'PERDIDA(CME)
      r_int_Cant = 0
      r_dbl_Monto = 0
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD = '003'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(14, 10), .Cells(14, 10)) = g_rst_Princi!CONT
            .Range(.Cells(14, 11), .Cells(14, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            .Range(.Cells(15, 10), .Cells(15, 10)) = r_int_Cant
            .Range(.Cells(15, 11), .Cells(15, 11)) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(16, 10), .Cells(16, 10)) = g_rst_Princi!CONT
            .Range(.Cells(16, 11), .Cells(16, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'PERDIDA(N.MIVIVIENDA)
      r_int_Cant = 0
      r_dbl_Monto = 0
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(18, 10), .Cells(18, 10)) = g_rst_Princi!CONT
            .Range(.Cells(18, 11), .Cells(18, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            .Range(.Cells(19, 10), .Cells(19, 10)) = r_int_Cant
            .Range(.Cells(19, 11), .Cells(19, 11)) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(20, 10), .Cells(20, 10)) = g_rst_Princi!CONT
            .Range(.Cells(20, 11), .Cells(20, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'PERDIDA(MICASAMAS)
      r_int_Cant = 0
      r_dbl_Monto = 0
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD = '019'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(22, 10), .Cells(22, 10)) = g_rst_Princi!CONT
            .Range(.Cells(22, 11), .Cells(22, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            .Range(.Cells(23, 10), .Cells(23, 10)) = r_int_Cant
            .Range(.Cells(23, 11), .Cells(23, 11)) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(24, 10), .Cells(24, 10)) = g_rst_Princi!CONT
            .Range(.Cells(24, 11), .Cells(24, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'PERDIDA(BBP)
      r_int_Cant = 0
      r_dbl_Monto = 0
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD IN ('021','022','023')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(26, 10), .Cells(26, 10)) = g_rst_Princi!CONT
            .Range(.Cells(26, 11), .Cells(26, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            .Range(.Cells(27, 10), .Cells(27, 10)) = r_int_Cant
            .Range(.Cells(27, 11), .Cells(27, 11)) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(28, 10), .Cells(28, 10)) = g_rst_Princi!CONT
            .Range(.Cells(28, 11), .Cells(28, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'PERDIDA(TECHO PROPIO)
      r_int_Cant = 0
      r_dbl_Monto = 0
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD = '024' "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .Range(.Cells(30, 10), .Cells(30, 10)) = g_rst_Princi!CONT
            .Range(.Cells(30, 11), .Cells(30, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            .Range(.Cells(31, 10), .Cells(31, 10)) = r_int_Cant
            .Range(.Cells(31, 11), .Cells(31, 11)) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .Range(.Cells(32, 10), .Cells(32, 10)) = g_rst_Princi!CONT
            .Range(.Cells(32, 11), .Cells(32, 11)) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      'DEFICIENTE
      .Cells(9, 4).Formula = "=D6+D7+D8"
      .Cells(9, 5).Formula = "=E6+E7+E8"
      .Cells(13, 4).Formula = "=D10+D11+D12"
      .Cells(13, 5).Formula = "=E10+E11+E12"
      .Cells(17, 4).Formula = "=D14+D15+D16"
      .Cells(17, 5).Formula = "=E14+E15+E16"
      .Cells(21, 4).Formula = "=D18+D19+D20"
      .Cells(21, 5).Formula = "=E18+E19+E20"
      .Cells(25, 4).Formula = "=D22+D23+D24"
      .Cells(25, 5).Formula = "=E22+E23+E24"
      .Cells(29, 4).Formula = "=D26+D27+D28"
      .Cells(29, 5).Formula = "=E26+E27+E28"
      .Cells(33, 4).Formula = "=D30+D31+D32"
      .Cells(33, 5).Formula = "=E30+E31+E32"
      
      'DUDOSO
      .Cells(9, 7).Formula = "=G6+G7+G8"
      .Cells(9, 8).Formula = "=H6+H7+H8"
      .Cells(13, 7).Formula = "=G10+G11+G12"
      .Cells(13, 8).Formula = "=H10+H11+H12"
      .Cells(17, 7).Formula = "=G14+G15+G16"
      .Cells(17, 8).Formula = "=H14+H15+H16"
      .Cells(21, 7).Formula = "=G18+G19+G20"
      .Cells(21, 8).Formula = "=H18+H19+H20"
      .Cells(25, 7).Formula = "=G22+G23+G24"
      .Cells(25, 8).Formula = "=H22+H23+H24"
      .Cells(29, 7).Formula = "=G26+G27+G28"
      .Cells(29, 8).Formula = "=H26+H27+H28"
      .Cells(33, 7).Formula = "=G30+G31+G32"
      .Cells(33, 8).Formula = "=H30+H31+H32"
      
      'PERDIDA
      .Cells(9, 10).Formula = "=J6+J7+J8"
      .Cells(9, 11).Formula = "=K6+K7+K8"
      .Cells(13, 10).Formula = "=J10+J11+J12"
      .Cells(13, 11).Formula = "=K10+K11+K12"
      .Cells(17, 10).Formula = "=J14+J15+J16"
      .Cells(17, 11).Formula = "=K14+K15+K16"
      .Cells(21, 10).Formula = "=J18+J19+J20"
      .Cells(21, 11).Formula = "=K18+K19+K20"
      .Cells(25, 10).Formula = "=J22+J23+J24"
      .Cells(25, 11).Formula = "=K22+K23+K24"
      .Cells(29, 10).Formula = "=J26+J27+J28"
      .Cells(29, 11).Formula = "=K26+K27+K28"
      .Cells(33, 10).Formula = "=J30+J31+J32"
      .Cells(33, 11).Formula = "=K30+K31+K32"
      
      'TOTALES
      .Cells(34, 4).Formula = "=D9+D13+D17+D21+D25+D29+D33"
      .Cells(34, 5).Formula = "=E9+E13+E17+E21+E25+E29+E33"
      .Cells(34, 7).Formula = "=G9+G13+G17+G21+G25+G29+G33"
      .Cells(34, 8).Formula = "=H9+H13+H17+H21+H25+H29+H33"
      .Cells(34, 10).Formula = "=J9+J13+J17+J21+J25+J29+J33"
      .Cells(34, 11).Formula = "=K9+K13+K17+K21+K25+K29+K33"
           
      For r_int_Contad = 6 To 34
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 3), .Cells(r_int_Contad, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
   End With
End Sub

Private Sub fs_Activa(ByVal estado As Boolean)
    cmd_ExpExc.Enabled = estado
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
   
   Call gs_LimpiaGrid(grd_LisCalif)
   
   tab_Clasif.Tab = 0
End Sub

Private Sub fs_Obtiene_Clasificacion()
   On Error Resume Next
   
   grd_LisCalif.Redraw = False
   Call gs_LimpiaGrid(grd_LisCalif)
      
   With grd_LisCalif
      .Rows = 21: .Cols = 12 '19
      .FixedRows = 4
      .MergeCells = flexMergeFree
      
      .TextMatrix(0, 0) = "PRODUCTO"
      .TextMatrix(1, 0) = "PRODUCTO"
      .TextMatrix(2, 0) = "PRODUCTO"
      .TextMatrix(3, 0) = "PRODUCTO"
      .TextMatrix(0, 1) = "CALIFICACION"
      .TextMatrix(0, 2) = "CALIFICACION"
      .TextMatrix(0, 3) = "CALIFICACION"
      .TextMatrix(0, 4) = "CALIFICACION"
      .TextMatrix(0, 5) = "CALIFICACION"
      .TextMatrix(0, 6) = "CALIFICACION"
      .TextMatrix(0, 7) = "CALIFICACION"
      .TextMatrix(0, 8) = "CALIFICACION"
      .TextMatrix(0, 9) = "CALIFICACION"
      .TextMatrix(0, 10) = "CALIFICACION"
      .TextMatrix(1, 1) = "NORMAL"
      .TextMatrix(1, 2) = "NORMAL"
      .TextMatrix(1, 3) = "CPP"
      .TextMatrix(1, 4) = "CPP"
      .TextMatrix(1, 5) = "DEFICIENTE"
      .TextMatrix(1, 6) = "DEFICIENTE"
      .TextMatrix(1, 7) = "DUDOSO"
      .TextMatrix(1, 8) = "DUDOSO"
      .TextMatrix(1, 9) = "PERDIDA"
      .TextMatrix(1, 10) = "PERDIDA"
      .TextMatrix(2, 1) = "01-30 DIAS"
      .TextMatrix(2, 2) = "01-30 DIAS"
      .TextMatrix(2, 3) = "31-60 DIAS"
      .TextMatrix(2, 4) = "31-60 DIAS"
      .TextMatrix(2, 5) = "61-120 DIAS"
      .TextMatrix(2, 6) = "61-120 DIAS"
      .TextMatrix(2, 7) = "121-365 DIAS"
      .TextMatrix(2, 8) = "121-365 DIAS"
      .TextMatrix(2, 9) = "MAS DE 365 DIAS"
      .TextMatrix(2, 10) = "MAS DE 365 DIAS"
      .TextMatrix(3, 1) = "N° CRED."
      .TextMatrix(3, 2) = "SALDO"
      .TextMatrix(3, 3) = "N° CRED."
      .TextMatrix(3, 4) = "SALDO"
      .TextMatrix(3, 5) = "N° CRED."
      .TextMatrix(3, 6) = "SALDO"
      .TextMatrix(3, 7) = "N° CRED."
      .TextMatrix(3, 8) = "SALDO"
      .TextMatrix(3, 9) = "N° CRED."
      .TextMatrix(3, 10) = "SALDO"
      .TextMatrix(0, 11) = "TOTAL"
      .TextMatrix(1, 11) = "TOTAL"
      .TextMatrix(2, 11) = "TOTAL"
      .TextMatrix(3, 11) = "TOTAL"
      .MergeRow(0) = True
      .MergeRow(1) = True
      .MergeRow(2) = True
      
      .MergeCol(0) = True
      .MergeCol(1) = True
      .MergeCol(2) = True
      .MergeCol(3) = True
      .MergeCol(4) = True
      .MergeCol(5) = True
      .MergeCol(6) = True
      .MergeCol(7) = True
      .MergeCol(8) = True
      .MergeCol(9) = True
      .MergeCol(10) = True
      .MergeCol(11) = True
            
      .TextMatrix(4, 0) = "CRC-PBP"
      .TextMatrix(5, 0) = "CRC-PBP"
      .TextMatrix(6, 0) = "MICASITA"
      .TextMatrix(7, 0) = "MICASITA"
      .TextMatrix(8, 0) = "CME"
      .TextMatrix(9, 0) = "CME"
      .TextMatrix(10, 0) = "N.MIVIVIENDA"
      .TextMatrix(11, 0) = "N.MIVIVIENDA"
      .TextMatrix(12, 0) = "MICASA MAS"
      .TextMatrix(13, 0) = "MICASA MAS"
      .TextMatrix(14, 0) = "BBP"
      .TextMatrix(15, 0) = "BBP"
      
      .TextMatrix(16, 0) = "TECHO PROPIO"
      .TextMatrix(17, 0) = "TECHO PROPIO"
      
      
      .TextMatrix(18, 0) = "PROMEDIO PONDERADO" '16
      .TextMatrix(19, 0) = "PROMEDIO PONDERADO"
      .TextMatrix(20, 0) = "TOTAL"
      
      .ColWidth(0) = 2000
      .ColWidth(1) = 900
      .ColWidth(2) = 1200
      .ColWidth(3) = 900
      .ColWidth(4) = 1200
      .ColWidth(5) = 900
      .ColWidth(6) = 1200
      .ColWidth(7) = 900
      .ColWidth(8) = 1200
      .ColWidth(9) = 900
      .ColWidth(10) = 1200
            
      .FixedAlignment(0) = flexAlignCenterCenter
      .FixedAlignment(1) = flexAlignCenterCenter
      .FixedAlignment(2) = flexAlignCenterCenter
      .FixedAlignment(3) = flexAlignCenterCenter
      .FixedAlignment(4) = flexAlignCenterCenter
      .FixedAlignment(5) = flexAlignCenterCenter
      .FixedAlignment(6) = flexAlignCenterCenter
      .FixedAlignment(7) = flexAlignCenterCenter
      .FixedAlignment(8) = flexAlignCenterCenter
      .FixedAlignment(9) = flexAlignCenterCenter
      .FixedAlignment(10) = flexAlignCenterCenter
      .FixedAlignment(11) = flexAlignCenterCenter

      'CRC-PBP
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CODPRD = '001'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .TextMatrix(4, 1) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(4, 2) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .TextMatrix(4, 3) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(4, 4) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .TextMatrix(4, 5) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(4, 6) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .TextMatrix(4, 7) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(4, 8) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .TextMatrix(4, 9) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(4, 10) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'MICASITA
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CODPRD IN ('002','006','011')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .TextMatrix(6, 1) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(6, 2) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .TextMatrix(6, 3) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(6, 4) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .TextMatrix(6, 5) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(6, 6) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .TextMatrix(6, 7) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(6, 8) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .TextMatrix(6, 9) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(6, 10) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'CME
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CODPRD = '003'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .TextMatrix(8, 1) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(8, 2) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .TextMatrix(8, 3) = Format(g_rst_Princi!CONT, "#,###")
            .TextMatrix(8, 4) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .TextMatrix(8, 5) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(8, 6) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .TextMatrix(8, 7) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(8, 8) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .TextMatrix(8, 9) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(8, 10) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
                   
         g_rst_Princi.MoveNext
      Loop

      'MIVIVIENDA
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .TextMatrix(10, 1) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(10, 2) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .TextMatrix(10, 3) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(10, 4) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .TextMatrix(10, 5) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(10, 6) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .TextMatrix(10, 7) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(10, 8) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .TextMatrix(10, 9) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(10, 10) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
                    
         g_rst_Princi.MoveNext
      Loop

      'MICASAMAS
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CODPRD = '019'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .TextMatrix(12, 1) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(12, 2) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .TextMatrix(12, 3) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(12, 4) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .TextMatrix(12, 5) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(12, 6) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .TextMatrix(12, 7) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(12, 8) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .TextMatrix(12, 9) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(12, 10) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'BBP
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CODPRD IN ('021','022','023')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .TextMatrix(14, 1) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(14, 2) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .TextMatrix(14, 3) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(14, 4) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .TextMatrix(14, 5) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(14, 6) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .TextMatrix(14, 7) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(14, 8) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .TextMatrix(14, 9) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(14, 10) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'TECHO PROPIO
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CODPRD = '024' "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
      g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            .TextMatrix(16, 1) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(16, 2) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
            .TextMatrix(16, 3) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(16, 4) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
            .TextMatrix(16, 5) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(16, 6) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
            .TextMatrix(16, 7) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(16, 8) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
            .TextMatrix(16, 9) = Format(g_rst_Princi!CONT, "#,##0")
            .TextMatrix(16, 10) = Format(g_rst_Princi!SALDO, "###,###,##0")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'Totales finales de la columna 11
      .TextMatrix(4, 11) = Format(Val(Format(.TextMatrix(4, 1), "#########")) + Val(Format(.TextMatrix(4, 3), "#########")) + Val(Format(.TextMatrix(4, 5), "#########")) + Val(Format(.TextMatrix(4, 7), "#########")) + Val(Format(.TextMatrix(4, 9), "#########")), "#,###,###")
      .TextMatrix(5, 11) = Format(Val(Format(.TextMatrix(4, 2), "#########")) + Val(Format(.TextMatrix(4, 4), "#########")) + Val(Format(.TextMatrix(4, 6), "#########")) + Val(Format(.TextMatrix(4, 8), "#########")) + Val(Format(.TextMatrix(4, 10), "#########")), "###,###,###,###")
      .TextMatrix(6, 11) = Format(Val(Format(.TextMatrix(6, 1), "#########")) + Val(Format(.TextMatrix(6, 3), "#########")) + Val(Format(.TextMatrix(6, 5), "#########")) + Val(Format(.TextMatrix(6, 7), "#########")) + Val(Format(.TextMatrix(6, 9), "#########")), "#,###,###")
      .TextMatrix(7, 11) = Format(Val(Format(.TextMatrix(6, 2), "#########")) + Val(Format(.TextMatrix(6, 4), "#########")) + Val(Format(.TextMatrix(6, 6), "#########")) + Val(Format(.TextMatrix(6, 8), "#########")) + Val(Format(.TextMatrix(6, 10), "#########")), "###,###,###,###")
      .TextMatrix(8, 11) = Format(Val(Format(.TextMatrix(8, 1), "#########")) + Val(Format(.TextMatrix(8, 3), "#########")) + Val(Format(.TextMatrix(8, 5), "#########")) + Val(Format(.TextMatrix(8, 7), "#########")) + Val(Format(.TextMatrix(8, 9), "#########")), "#,###,###")
      .TextMatrix(9, 11) = Format(Val(Format(.TextMatrix(8, 2), "#########")) + Val(Format(.TextMatrix(8, 4), "#########")) + Val(Format(.TextMatrix(8, 6), "#########")) + Val(Format(.TextMatrix(8, 8), "#########")) + Val(Format(.TextMatrix(8, 10), "#########")), "###,###,###,###")
      .TextMatrix(10, 11) = Format(Val(Format(.TextMatrix(10, 1), "#########")) + Val(Format(.TextMatrix(10, 3), "#########")) + Val(Format(.TextMatrix(10, 5), "#########")) + Val(Format(.TextMatrix(10, 7), "#########")) + Val(Format(.TextMatrix(10, 9), "#########")), "#,###,###")
      .TextMatrix(11, 11) = Format(Val(Format(.TextMatrix(10, 2), "#########")) + Val(Format(.TextMatrix(10, 4), "#########")) + Val(Format(.TextMatrix(10, 6), "#########")) + Val(Format(.TextMatrix(10, 8), "#########")) + Val(Format(.TextMatrix(10, 10), "#########")), "###,###,###,###")
      .TextMatrix(12, 11) = Format(Val(Format(.TextMatrix(12, 1), "#########")) + Val(Format(.TextMatrix(12, 3), "#########")) + Val(Format(.TextMatrix(12, 5), "#########")) + Val(Format(.TextMatrix(12, 7), "#########")) + Val(Format(.TextMatrix(12, 9), "#########")), "#,###,###")
      .TextMatrix(13, 11) = Format(Val(Format(.TextMatrix(12, 2), "#########")) + Val(Format(.TextMatrix(12, 4), "#########")) + Val(Format(.TextMatrix(12, 6), "#########")) + Val(Format(.TextMatrix(12, 8), "#########")) + Val(Format(.TextMatrix(12, 10), "#########")), "###,###,###,###")
      .TextMatrix(14, 11) = Format(Val(Format(.TextMatrix(14, 1), "#########")) + Val(Format(.TextMatrix(14, 3), "#########")) + Val(Format(.TextMatrix(14, 5), "#########")) + Val(Format(.TextMatrix(14, 7), "#########")) + Val(Format(.TextMatrix(14, 9), "#########")), "#,###,###")
      .TextMatrix(15, 11) = Format(Val(Format(.TextMatrix(14, 2), "#########")) + Val(Format(.TextMatrix(14, 4), "#########")) + Val(Format(.TextMatrix(14, 6), "#########")) + Val(Format(.TextMatrix(14, 8), "#########")) + Val(Format(.TextMatrix(14, 10), "#########")), "###,###,###,###")
      .TextMatrix(16, 11) = Format(Val(Format(.TextMatrix(16, 1), "#########")) + Val(Format(.TextMatrix(16, 3), "#########")) + Val(Format(.TextMatrix(16, 5), "#########")) + Val(Format(.TextMatrix(16, 7), "#########")) + Val(Format(.TextMatrix(16, 9), "#########")), "#,###,###")
      .TextMatrix(17, 11) = Format(Val(Format(.TextMatrix(16, 2), "#########")) + Val(Format(.TextMatrix(16, 4), "#########")) + Val(Format(.TextMatrix(16, 6), "#########")) + Val(Format(.TextMatrix(16, 8), "#########")) + Val(Format(.TextMatrix(16, 10), "#########")), "###,###,###,###")
      .TextMatrix(18, 11) = Format(Val(Format(.TextMatrix(4, 11), "#########")) + Val(Format(.TextMatrix(6, 11), "#########")) + Val(Format(.TextMatrix(8, 11), "#########")) + Val(Format(.TextMatrix(10, 11), "#########")) + Val(Format(.TextMatrix(12, 11), "#########")) + Val(Format(.TextMatrix(14, 11), "#########")) + Val(Format(.TextMatrix(16, 11), "#########")), "#,###,###")
      .TextMatrix(19, 11) = Format(Val(Format(.TextMatrix(5, 11), "#########")) + Val(Format(.TextMatrix(7, 11), "#########")) + Val(Format(.TextMatrix(9, 11), "#########")) + Val(Format(.TextMatrix(11, 11), "#########")) + Val(Format(.TextMatrix(13, 11), "#########")) + Val(Format(.TextMatrix(15, 11), "#########")) + Val(Format(.TextMatrix(17, 11), "#########")), "#,###,###")

      'Porcentajes
      .TextMatrix(5, 1) = Format((Val(Format(.TextMatrix(4, 1), "#########")) / .TextMatrix(4, 11)) * 100, "##0.00") & "%"
      .TextMatrix(5, 2) = Format((Val(Format(.TextMatrix(4, 2), "#########")) / .TextMatrix(5, 11)) * 100, "##0.00") & "%"
      .TextMatrix(5, 3) = Format((Val(Format(.TextMatrix(4, 3), "#########")) / .TextMatrix(4, 11)) * 100, "##0.00") & "%"
      .TextMatrix(5, 4) = Format((Val(Format(.TextMatrix(4, 4), "#########")) / .TextMatrix(5, 11)) * 100, "##0.00") & "%"
      .TextMatrix(5, 5) = Format((Val(Format(.TextMatrix(4, 5), "#########")) / .TextMatrix(4, 11)) * 100, "##0.00") & "%"
      .TextMatrix(5, 6) = Format((Val(Format(.TextMatrix(4, 6), "#########")) / .TextMatrix(5, 11)) * 100, "##0.00") & "%"
      .TextMatrix(5, 7) = Format((Val(Format(.TextMatrix(4, 7), "#########")) / .TextMatrix(4, 11)) * 100, "##0.00") & "%"
      .TextMatrix(5, 8) = Format((Val(Format(.TextMatrix(4, 8), "#########")) / .TextMatrix(5, 11)) * 100, "##0.00") & "%"
      .TextMatrix(5, 9) = Format((Val(Format(.TextMatrix(4, 9), "#########")) / .TextMatrix(4, 11)) * 100, "##0.00") & "%"
      .TextMatrix(5, 10) = Format((Val(Format(.TextMatrix(4, 10), "#########")) / .TextMatrix(5, 11)) * 100, "##0.00") & "%"
      
      .TextMatrix(7, 1) = Format((Val(Format(.TextMatrix(6, 1), "#########")) / .TextMatrix(6, 11)) * 100, "##0.00") & "%"
      .TextMatrix(7, 2) = Format((Val(Format(.TextMatrix(6, 2), "#########")) / .TextMatrix(7, 11)) * 100, "##0.00") & "%"
      .TextMatrix(7, 3) = Format((Val(Format(.TextMatrix(6, 3), "#########")) / .TextMatrix(6, 11)) * 100, "##0.00") & "%"
      .TextMatrix(7, 4) = Format((Val(Format(.TextMatrix(6, 4), "#########")) / .TextMatrix(7, 11)) * 100, "##0.00") & "%"
      .TextMatrix(7, 5) = Format((Val(Format(.TextMatrix(6, 5), "#########")) / .TextMatrix(6, 11)) * 100, "##0.00") & "%"
      .TextMatrix(7, 6) = Format((Val(Format(.TextMatrix(6, 6), "#########")) / .TextMatrix(7, 11)) * 100, "##0.00") & "%"
      .TextMatrix(7, 7) = Format((Val(Format(.TextMatrix(6, 7), "#########")) / .TextMatrix(6, 11)) * 100, "##0.00") & "%"
      .TextMatrix(7, 8) = Format((Val(Format(.TextMatrix(6, 8), "#########")) / .TextMatrix(7, 11)) * 100, "##0.00") & "%"
      .TextMatrix(7, 9) = Format((Val(Format(.TextMatrix(6, 9), "#########")) / .TextMatrix(6, 11)) * 100, "##0.00") & "%"
      .TextMatrix(7, 10) = Format((Val(Format(.TextMatrix(6, 10), "#########")) / .TextMatrix(7, 11)) * 100, "##0.00") & "%"
      
      .TextMatrix(9, 1) = Format((Val(Format(.TextMatrix(8, 1), "#########")) / .TextMatrix(8, 11)) * 100, "##0.00") & "%"
      .TextMatrix(9, 2) = Format((Val(Format(.TextMatrix(8, 2), "#########")) / .TextMatrix(9, 11)) * 100, "##0.00") & "%"
      .TextMatrix(9, 3) = Format((Val(Format(.TextMatrix(8, 3), "#########")) / .TextMatrix(8, 11)) * 100, "##0.00") & "%"
      .TextMatrix(9, 4) = Format((Val(Format(.TextMatrix(8, 4), "#########")) / .TextMatrix(9, 11)) * 100, "##0.00") & "%"
      .TextMatrix(9, 5) = Format((Val(Format(.TextMatrix(8, 5), "#########")) / .TextMatrix(8, 11)) * 100, "##0.00") & "%"
      .TextMatrix(9, 6) = Format((Val(Format(.TextMatrix(8, 6), "#########")) / .TextMatrix(9, 11)) * 100, "##0.00") & "%"
      .TextMatrix(9, 7) = Format((Val(Format(.TextMatrix(8, 7), "#########")) / .TextMatrix(8, 11)) * 100, "##0.00") & "%"
      .TextMatrix(9, 8) = Format((Val(Format(.TextMatrix(8, 8), "#########")) / .TextMatrix(9, 11)) * 100, "##0.00") & "%"
      .TextMatrix(9, 9) = Format((Val(Format(.TextMatrix(8, 9), "#########")) / .TextMatrix(8, 11)) * 100, "##0.00") & "%"
      .TextMatrix(9, 10) = Format((Val(Format(.TextMatrix(8, 10), "#########")) / .TextMatrix(9, 11)) * 100, "##0.00") & "%"
      
      .TextMatrix(11, 1) = Format((Val(Format(.TextMatrix(10, 1), "#########")) / .TextMatrix(10, 11)) * 100, "##0.00") & "%"
      .TextMatrix(11, 2) = Format((Val(Format(.TextMatrix(10, 2), "#########")) / .TextMatrix(11, 11)) * 100, "##0.00") & "%"
      .TextMatrix(11, 3) = Format((Val(Format(.TextMatrix(10, 3), "#########")) / .TextMatrix(10, 11)) * 100, "##0.00") & "%"
      .TextMatrix(11, 4) = Format((Val(Format(.TextMatrix(10, 4), "#########")) / .TextMatrix(11, 11)) * 100, "##0.00") & "%"
      .TextMatrix(11, 5) = Format((Val(Format(.TextMatrix(10, 5), "#########")) / .TextMatrix(10, 11)) * 100, "##0.00") & "%"
      .TextMatrix(11, 6) = Format((Val(Format(.TextMatrix(10, 6), "#########")) / .TextMatrix(11, 11)) * 100, "##0.00") & "%"
      .TextMatrix(11, 7) = Format((Val(Format(.TextMatrix(10, 7), "#########")) / .TextMatrix(10, 11)) * 100, "##0.00") & "%"
      .TextMatrix(11, 8) = Format((Val(Format(.TextMatrix(10, 8), "#########")) / .TextMatrix(11, 11)) * 100, "##0.00") & "%"
      .TextMatrix(11, 9) = Format((Val(Format(.TextMatrix(10, 9), "#########")) / .TextMatrix(10, 11)) * 100, "##0.00") & "%"
      .TextMatrix(11, 10) = Format((Val(Format(.TextMatrix(10, 10), "#########")) / .TextMatrix(11, 11)) * 100, "##0.00") & "%"
      
      .TextMatrix(13, 1) = Format((Val(Format(.TextMatrix(12, 1), "#########")) / .TextMatrix(12, 11)) * 100, "##0.00") & "%"
      .TextMatrix(13, 2) = Format((Val(Format(.TextMatrix(12, 2), "#########")) / .TextMatrix(13, 11)) * 100, "##0.00") & "%"
      .TextMatrix(13, 3) = Format((Val(Format(.TextMatrix(12, 3), "#########")) / .TextMatrix(12, 11)) * 100, "##0.00") & "%"
      .TextMatrix(13, 4) = Format((Val(Format(.TextMatrix(12, 4), "#########")) / .TextMatrix(13, 11)) * 100, "##0.00") & "%"
      .TextMatrix(13, 5) = Format((Val(Format(.TextMatrix(12, 5), "#########")) / .TextMatrix(12, 11)) * 100, "##0.00") & "%"
      .TextMatrix(13, 6) = Format((Val(Format(.TextMatrix(12, 6), "#########")) / .TextMatrix(13, 11)) * 100, "##0.00") & "%"
      .TextMatrix(13, 7) = Format((Val(Format(.TextMatrix(12, 7), "#########")) / .TextMatrix(12, 11)) * 100, "##0.00") & "%"
      .TextMatrix(13, 8) = Format((Val(Format(.TextMatrix(12, 8), "#########")) / .TextMatrix(13, 11)) * 100, "##0.00") & "%"
      .TextMatrix(13, 9) = Format((Val(Format(.TextMatrix(12, 9), "#########")) / .TextMatrix(12, 11)) * 100, "##0.00") & "%"
      .TextMatrix(13, 10) = Format((Val(Format(.TextMatrix(12, 10), "#########")) / .TextMatrix(13, 11)) * 100, "##0.00") & "%"

      .TextMatrix(15, 1) = Format((Val(Format(.TextMatrix(14, 1), "#########")) / .TextMatrix(14, 11)) * 100, "##0.00") & "%"
      .TextMatrix(15, 2) = Format((Val(Format(.TextMatrix(14, 2), "#########")) / .TextMatrix(15, 11)) * 100, "##0.00") & "%"
      .TextMatrix(15, 3) = Format((Val(Format(.TextMatrix(14, 3), "#########")) / .TextMatrix(14, 11)) * 100, "##0.00") & "%"
      .TextMatrix(15, 4) = Format((Val(Format(.TextMatrix(14, 4), "#########")) / .TextMatrix(15, 11)) * 100, "##0.00") & "%"
      .TextMatrix(15, 5) = Format((Val(Format(.TextMatrix(14, 5), "#########")) / .TextMatrix(14, 11)) * 100, "##0.00") & "%"
      .TextMatrix(15, 6) = Format((Val(Format(.TextMatrix(14, 6), "#########")) / .TextMatrix(15, 11)) * 100, "##0.00") & "%"
      .TextMatrix(15, 7) = Format((Val(Format(.TextMatrix(14, 7), "#########")) / .TextMatrix(14, 11)) * 100, "##0.00") & "%"
      .TextMatrix(15, 8) = Format((Val(Format(.TextMatrix(14, 8), "#########")) / .TextMatrix(15, 11)) * 100, "##0.00") & "%"
      .TextMatrix(15, 9) = Format((Val(Format(.TextMatrix(14, 9), "#########")) / .TextMatrix(14, 11)) * 100, "##0.00") & "%"
      .TextMatrix(15, 10) = Format((Val(Format(.TextMatrix(14, 10), "#########")) / .TextMatrix(15, 11)) * 100, "##0.00") & "%"
      
      .TextMatrix(17, 1) = Format((Val(Format(.TextMatrix(16, 1), "#########")) / .TextMatrix(16, 11)) * 100, "##0.00") & "%"
      .TextMatrix(17, 2) = Format((Val(Format(.TextMatrix(16, 2), "#########")) / .TextMatrix(17, 11)) * 100, "##0.00") & "%"
      .TextMatrix(17, 3) = Format((Val(Format(.TextMatrix(16, 3), "#########")) / .TextMatrix(16, 11)) * 100, "##0.00") & "%"
      .TextMatrix(17, 4) = Format((Val(Format(.TextMatrix(16, 4), "#########")) / .TextMatrix(17, 11)) * 100, "##0.00") & "%"
      .TextMatrix(17, 5) = Format((Val(Format(.TextMatrix(16, 5), "#########")) / .TextMatrix(16, 11)) * 100, "##0.00") & "%"
      .TextMatrix(17, 6) = Format((Val(Format(.TextMatrix(16, 6), "#########")) / .TextMatrix(17, 11)) * 100, "##0.00") & "%"
      .TextMatrix(17, 7) = Format((Val(Format(.TextMatrix(16, 7), "#########")) / .TextMatrix(16, 11)) * 100, "##0.00") & "%"
      .TextMatrix(17, 8) = Format((Val(Format(.TextMatrix(16, 8), "#########")) / .TextMatrix(17, 11)) * 100, "##0.00") & "%"
      .TextMatrix(17, 9) = Format((Val(Format(.TextMatrix(16, 9), "#########")) / .TextMatrix(16, 11)) * 100, "##0.00") & "%"
      .TextMatrix(17, 10) = Format((Val(Format(.TextMatrix(16, 10), "#########")) / .TextMatrix(17, 11)) * 100, "##0.00") & "%"
      
      'Suma Final para Total
      .TextMatrix(20, 1) = Format(Val(Format(.TextMatrix(4, 1), "#########")) + Val(Format(.TextMatrix(6, 1), "#########")) + Val(Format(.TextMatrix(8, 1), "#########")) + Val(Format(.TextMatrix(10, 1), "#########")) + Val(Format(.TextMatrix(12, 1), "#########")) + Val(Format(.TextMatrix(14, 1), "#########")) + Val(Format(.TextMatrix(16, 1), "#########")), "#,###,###")
      .TextMatrix(20, 2) = Format(Val(Format(.TextMatrix(4, 2), "#########")) + Val(Format(.TextMatrix(6, 2), "#########")) + Val(Format(.TextMatrix(8, 2), "#########")) + Val(Format(.TextMatrix(10, 2), "#########")) + Val(Format(.TextMatrix(12, 2), "#########")) + Val(Format(.TextMatrix(14, 2), "#########")) + Val(Format(.TextMatrix(16, 2), "#########")), "###,###,###,###")
      .TextMatrix(20, 3) = Format(Val(Format(.TextMatrix(4, 3), "#########")) + Val(Format(.TextMatrix(6, 3), "#########")) + Val(Format(.TextMatrix(8, 3), "#########")) + Val(Format(.TextMatrix(10, 3), "#########")) + Val(Format(.TextMatrix(12, 3), "#########")) + Val(Format(.TextMatrix(14, 3), "#########")) + Val(Format(.TextMatrix(16, 3), "#########")), "#,###,###")
      .TextMatrix(20, 4) = Format(Val(Format(.TextMatrix(4, 4), "#########")) + Val(Format(.TextMatrix(6, 4), "#########")) + Val(Format(.TextMatrix(8, 4), "#########")) + Val(Format(.TextMatrix(10, 4), "#########")) + Val(Format(.TextMatrix(12, 4), "#########")) + Val(Format(.TextMatrix(14, 4), "#########")) + Val(Format(.TextMatrix(16, 4), "#########")), "###,###,###,###")
      .TextMatrix(20, 5) = Format(Val(Format(.TextMatrix(4, 5), "#########")) + Val(Format(.TextMatrix(6, 5), "#########")) + Val(Format(.TextMatrix(8, 5), "#########")) + Val(Format(.TextMatrix(10, 5), "#########")) + Val(Format(.TextMatrix(12, 5), "#########")) + Val(Format(.TextMatrix(14, 5), "#########")) + Val(Format(.TextMatrix(16, 5), "#########")), "#,###,###")
      .TextMatrix(20, 6) = Format(Val(Format(.TextMatrix(4, 6), "#########")) + Val(Format(.TextMatrix(6, 6), "#########")) + Val(Format(.TextMatrix(8, 6), "#########")) + Val(Format(.TextMatrix(10, 6), "#########")) + Val(Format(.TextMatrix(12, 6), "#########")) + Val(Format(.TextMatrix(14, 6), "#########")) + Val(Format(.TextMatrix(16, 6), "#########")), "###,###,###,###")
      .TextMatrix(20, 7) = Format(Val(Format(.TextMatrix(4, 7), "#########")) + Val(Format(.TextMatrix(6, 7), "#########")) + Val(Format(.TextMatrix(8, 7), "#########")) + Val(Format(.TextMatrix(10, 7), "#########")) + Val(Format(.TextMatrix(12, 7), "#########")) + Val(Format(.TextMatrix(14, 7), "#########")) + Val(Format(.TextMatrix(16, 7), "#########")), "#,###,###")
      .TextMatrix(20, 8) = Format(Val(Format(.TextMatrix(4, 8), "#########")) + Val(Format(.TextMatrix(6, 8), "#########")) + Val(Format(.TextMatrix(8, 8), "#########")) + Val(Format(.TextMatrix(10, 8), "#########")) + Val(Format(.TextMatrix(12, 8), "#########")) + Val(Format(.TextMatrix(14, 8), "#########")) + Val(Format(.TextMatrix(16, 8), "#########")), "###,###,###,###")
      .TextMatrix(20, 9) = Format(Val(Format(.TextMatrix(4, 9), "#########")) + Val(Format(.TextMatrix(6, 9), "#########")) + Val(Format(.TextMatrix(8, 9), "#########")) + Val(Format(.TextMatrix(10, 9), "#########")) + Val(Format(.TextMatrix(12, 9), "#########")) + Val(Format(.TextMatrix(14, 9), "#########")) + Val(Format(.TextMatrix(16, 9), "#########")), "#,###,###")
      .TextMatrix(20, 10) = Format(Val(Format(.TextMatrix(4, 10), "#########")) + Val(Format(.TextMatrix(6, 10), "#########")) + Val(Format(.TextMatrix(8, 10), "#########")) + Val(Format(.TextMatrix(10, 10), "#########")) + Val(Format(.TextMatrix(12, 10), "#########")) + Val(Format(.TextMatrix(14, 10), "#########")) + Val(Format(.TextMatrix(16, 10), "#########")), "###,###,###,###")

      'Porcentaje Promedio Ponderado (abajo repetimos la misma rutina para mostrar en una sola fila (* MERGE))
      .TextMatrix(18, 1) = Format((Val(Format(.TextMatrix(20, 1), "#########")) / Val(Format(.TextMatrix(18, 11), "#########"))) * 100, "##0.00") & "%" '16
      .TextMatrix(18, 2) = Format((Val(Format(.TextMatrix(20, 2), "#########")) / Val(Format(.TextMatrix(19, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(18, 3) = Format((Val(Format(.TextMatrix(20, 3), "#########")) / Val(Format(.TextMatrix(18, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(18, 4) = Format((Val(Format(.TextMatrix(20, 4), "#########")) / Val(Format(.TextMatrix(19, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(18, 5) = Format((Val(Format(.TextMatrix(20, 5), "#########")) / Val(Format(.TextMatrix(18, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(18, 6) = Format((Val(Format(.TextMatrix(20, 6), "#########")) / Val(Format(.TextMatrix(19, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(18, 7) = Format((Val(Format(.TextMatrix(20, 7), "#########")) / Val(Format(.TextMatrix(18, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(18, 8) = Format((Val(Format(.TextMatrix(20, 8), "#########")) / Val(Format(.TextMatrix(19, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(18, 9) = Format((Val(Format(.TextMatrix(20, 9), "#########")) / Val(Format(.TextMatrix(18, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(18, 10) = Format((Val(Format(.TextMatrix(20, 10), "#########")) / Val(Format(.TextMatrix(19, 11), "#########"))) * 100, "##0.00") & "%"
      '(*)
      .TextMatrix(19, 1) = Format((Val(Format(.TextMatrix(20, 1), "#########")) / Val(Format(.TextMatrix(18, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(19, 2) = Format((Val(Format(.TextMatrix(20, 2), "#########")) / Val(Format(.TextMatrix(19, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(19, 3) = Format((Val(Format(.TextMatrix(20, 3), "#########")) / Val(Format(.TextMatrix(18, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(19, 4) = Format((Val(Format(.TextMatrix(20, 4), "#########")) / Val(Format(.TextMatrix(19, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(19, 5) = Format((Val(Format(.TextMatrix(20, 5), "#########")) / Val(Format(.TextMatrix(18, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(19, 6) = Format((Val(Format(.TextMatrix(20, 6), "#########")) / Val(Format(.TextMatrix(19, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(19, 7) = Format((Val(Format(.TextMatrix(20, 7), "#########")) / Val(Format(.TextMatrix(18, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(19, 8) = Format((Val(Format(.TextMatrix(20, 8), "#########")) / Val(Format(.TextMatrix(19, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(19, 9) = Format((Val(Format(.TextMatrix(20, 9), "#########")) / Val(Format(.TextMatrix(18, 11), "#########"))) * 100, "##0.00") & "%"
      .TextMatrix(19, 10) = Format((Val(Format(.TextMatrix(20, 10), "#########")) / Val(Format(.TextMatrix(19, 11), "#########"))) * 100, "##0.00") & "%"
            
      .Col = 0: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 0: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 1: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 1: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 10: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 10: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 11: .Row = 18: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 11: .Row = 19: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      
      .Col = 0: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 1: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 10: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 11: .Row = 20: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
   End With
  
   grd_LisCalif.Redraw = True
   grd_LisCalif.Enabled = True
   Call gs_UbicaGrid(grd_LisCalif, 2)
End Sub

Private Sub fs_Obtiene_Detalle_Clasificacion()
Dim r_int_fila       As Integer
Dim r_Cad_Ante       As String
Dim r_dbl_Suma1      As Double
Dim r_dbl_Suma2      As Double
Dim r_dbl_Suma3      As Double
Dim r_dbl_Saldo1     As Double
Dim r_dbl_Saldo2     As Double
Dim r_dbl_Saldo3     As Double

   grd_LisDetCalif.Redraw = False
   Call gs_LimpiaGrid(grd_LisDetCalif)
      
   With grd_LisDetCalif
      .Rows = 13: .Cols = 6: .RowHeight(0) = 400
      .FixedRows = 1
      .MergeCells = flexMergeFree
      
      .TextMatrix(0, 0) = "ESTADO"
      .TextMatrix(0, 1) = "SITUACION"
      .TextMatrix(0, 2) = "N° CRED."
      .TextMatrix(0, 3) = "%"
      .TextMatrix(0, 4) = "SALDO"
      .TextMatrix(0, 5) = "%"
                  
      .MergeCol(0) = True
            
      .TextMatrix(1, 0) = "DEFICIENTE"
      .TextMatrix(2, 0) = "DEFICIENTE"
      .TextMatrix(3, 0) = "DEFICIENTE"
      .TextMatrix(4, 0) = "DEFICIENTE"
      .TextMatrix(5, 0) = "DUDOSO"
      .TextMatrix(6, 0) = "DUDOSO"
      .TextMatrix(7, 0) = "DUDOSO"
      .TextMatrix(8, 0) = "DUDOSO"
      .TextMatrix(9, 0) = "PERDIDA"
      .TextMatrix(10, 0) = "PERDIDA"
      .TextMatrix(11, 0) = "PERDIDA"
      .TextMatrix(12, 0) = "PERDIDA"
      
      .TextMatrix(1, 1) = "ALINEADO"
      .TextMatrix(2, 1) = "ALINEADO MOROSO"
      .TextMatrix(3, 1) = "MOROSO"
      .TextMatrix(4, 1) = "TOTAL"
      .TextMatrix(5, 1) = "ALINEADO"
      .TextMatrix(6, 1) = "ALINEADO MOROSO"
      .TextMatrix(7, 1) = "MOROSO"
      .TextMatrix(8, 1) = "TOTAL"
      .TextMatrix(9, 1) = "ALINEADO"
      .TextMatrix(10, 1) = "ALINEADO MOROSO"
      .TextMatrix(11, 1) = "MOROSO"
      .TextMatrix(12, 1) = "TOTAL"
            
      .ColWidth(0) = 3000
      .ColWidth(1) = 4000
      .ColWidth(2) = 1200
      .ColWidth(3) = 1200
      .ColWidth(4) = 2500
      .ColWidth(5) = 1200
            
      .FixedAlignment(0) = flexAlignCenterCenter
      .FixedAlignment(1) = flexAlignCenterCenter
      .FixedAlignment(2) = flexAlignCenterCenter
      .FixedAlignment(3) = flexAlignCenterCenter
      .FixedAlignment(4) = flexAlignCenterCenter
      .FixedAlignment(5) = flexAlignCenterCenter
      
      .ColAlignment(0) = flexAlignCenterCenter
      .Col = 1: .Row = 4: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 4: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 4: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 4: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 4: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 1: .Row = 8: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 8: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 8: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 8: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 8: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 1: .Row = 12: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 12: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 12: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 12: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 12: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
 
      'Procesa infomacion
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "USP_RPT_CLASIFICACION("
      g_str_Parame = g_str_Parame & CInt(r_int_PerAno) & ", "
      g_str_Parame = g_str_Parame & CInt(r_int_PerMes) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'REPORTE_CLASIFICACION', "
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         grd_LisDetCalif.Redraw = True
         MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      r_int_fila = 1
      Do While Not g_rst_Princi.EOF
         r_Cad_Ante = g_rst_Princi!RPT_VALCAD01
         If g_rst_Princi!RPT_VALCAD01 = 2 Then
            .TextMatrix(r_int_fila, 2) = Format(g_rst_Princi!RPT_VALNUM01, "#,##0")
            .TextMatrix(r_int_fila, 3) = IIf(IsNull(g_rst_Princi!RPT_VALNUM02), "", Format(g_rst_Princi!RPT_VALNUM02, "##0.00") & "%")
            .TextMatrix(r_int_fila, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,##0.00")
            .TextMatrix(r_int_fila, 5) = IIf(IsNull(g_rst_Princi!RPT_VALNUM04), "", Format(g_rst_Princi!RPT_VALNUM04, "##0.00") & "%")
            
            .TextMatrix(4, 2) = Val(.TextMatrix(4, 2)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01)
            .TextMatrix(4, 3) = Val(.TextMatrix(4, 3)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02)
            .TextMatrix(4, 4) = Val(.TextMatrix(4, 4)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03)
            .TextMatrix(4, 5) = Val(.TextMatrix(4, 5)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04)
            
            r_dbl_Suma1 = r_dbl_Suma1 + IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02)
            Select Case r_dbl_Suma1
               Case 99.99: .TextMatrix(r_int_fila, 3) = Format(g_rst_Princi!RPT_VALNUM02 + 0.01, "##0.00") & "%"
               Case 100.01: .TextMatrix(r_int_fila, 3) = Format(g_rst_Princi!RPT_VALNUM02 - 0.01, "##0.00") & "%"
            End Select
            
            r_dbl_Saldo1 = r_dbl_Saldo1 + IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04)
            Select Case r_dbl_Saldo1
               Case 99.99: .TextMatrix(r_int_fila, 5) = Format(g_rst_Princi!RPT_VALNUM04 + 0.01, "##0.00") & "%"
               Case 100.01: .TextMatrix(r_int_fila, 5) = Format(g_rst_Princi!RPT_VALNUM04 - 0.01, "##0.00") & "%"
            End Select
         End If
         If g_rst_Princi!RPT_VALCAD01 = 3 Then
            .TextMatrix(r_int_fila, 2) = Format(g_rst_Princi!RPT_VALNUM01, "#,##0")
            .TextMatrix(r_int_fila, 3) = IIf(IsNull(g_rst_Princi!RPT_VALNUM02), "", Format(g_rst_Princi!RPT_VALNUM02, "##0.00") & "%")
            .TextMatrix(r_int_fila, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,##0.00")
            .TextMatrix(r_int_fila, 5) = IIf(IsNull(g_rst_Princi!RPT_VALNUM04), "", Format(g_rst_Princi!RPT_VALNUM04, "##0.00") & "%")
            
            .TextMatrix(8, 2) = Val(.TextMatrix(8, 2)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01)
            .TextMatrix(8, 3) = Val(.TextMatrix(8, 3)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02)
            .TextMatrix(8, 4) = Val(.TextMatrix(8, 4)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03)
            .TextMatrix(8, 5) = Val(.TextMatrix(8, 5)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04)
            
            r_dbl_Suma2 = r_dbl_Suma2 + IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02)
            Select Case r_dbl_Suma2
               Case 99.99: .TextMatrix(r_int_fila, 3) = Format(g_rst_Princi!RPT_VALNUM02 + 0.01, "##0.00") & "%"
               Case 100.01: .TextMatrix(r_int_fila, 3) = Format(g_rst_Princi!RPT_VALNUM02 - 0.01, "##0.00") & "%"
            End Select
            
            r_dbl_Saldo2 = r_dbl_Saldo2 + IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04)
            Select Case r_dbl_Saldo2
               Case 99.99: .TextMatrix(r_int_fila, 5) = Format(g_rst_Princi!RPT_VALNUM04 + 0.01, "##0.00") & "%"
               Case 100.01: .TextMatrix(r_int_fila, 5) = Format(g_rst_Princi!RPT_VALNUM04 - 0.01, "##0.00") & "%"
            End Select
         End If
         If g_rst_Princi!RPT_VALCAD01 = 4 Then
            .TextMatrix(r_int_fila, 2) = Format(g_rst_Princi!RPT_VALNUM01, "#,##0")
            .TextMatrix(r_int_fila, 3) = IIf(IsNull(g_rst_Princi!RPT_VALNUM02), "", Format(g_rst_Princi!RPT_VALNUM02, "##0.00") & "%")
            .TextMatrix(r_int_fila, 4) = Format(g_rst_Princi!RPT_VALNUM03, "###,##0.00")
            .TextMatrix(r_int_fila, 5) = IIf(IsNull(g_rst_Princi!RPT_VALNUM04), "", Format(g_rst_Princi!RPT_VALNUM04, "##0.00") & "%")
            
            .TextMatrix(12, 2) = Val(.TextMatrix(12, 2)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01)
            .TextMatrix(12, 3) = Val(.TextMatrix(12, 3)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02)
            .TextMatrix(12, 4) = Val(.TextMatrix(12, 4)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03)
            .TextMatrix(12, 5) = Val(.TextMatrix(12, 5)) + IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04)

            r_dbl_Suma3 = r_dbl_Suma3 + IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02)
            Select Case r_dbl_Suma3
               Case 99.99: .TextMatrix(r_int_fila, 3) = Format(g_rst_Princi!RPT_VALNUM02 + 0.01, "##0.00") & "%"
               Case 100.01: .TextMatrix(r_int_fila, 3) = Format(g_rst_Princi!RPT_VALNUM02 - 0.01, "##0.00") & "%"
            End Select
            
            r_dbl_Saldo3 = r_dbl_Saldo3 + IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04)
            Select Case r_dbl_Saldo3
               Case 99.99: .TextMatrix(r_int_fila, 5) = Format(g_rst_Princi!RPT_VALNUM04 + 0.01, "##0.00") & "%"
               Case 100.01: .TextMatrix(r_int_fila, 5) = Format(g_rst_Princi!RPT_VALNUM04 - 0.01, "##0.00") & "%"
            End Select
         End If
         r_int_fila = r_int_fila + 1
         g_rst_Princi.MoveNext
         
         If Not g_rst_Princi.EOF Then
            If r_Cad_Ante <> g_rst_Princi!RPT_VALCAD01 Then
               r_int_fila = r_int_fila + 1
            End If
         End If
      Loop
      
      'DEFICIENTE
      If .TextMatrix(4, 3) <> 0 Then
         If .TextMatrix(4, 3) < 100 Or .TextMatrix(4, 3) > 100 Then
            .TextMatrix(4, 3) = "100" & "%"
         Else
            .TextMatrix(4, 3) = "100" & "%"
         End If
      End If

      If .TextMatrix(4, 5) <> 0 Then
         If .TextMatrix(4, 5) < 100 Or .TextMatrix(4, 5) > 100 Then
            .TextMatrix(4, 5) = "100" & "%"
         Else
            .TextMatrix(4, 5) = "100" & "%"
         End If
      End If
      .TextMatrix(4, 4) = Format(.TextMatrix(4, 4), "###,##0.00")
      
      'DUDOSO
      If .TextMatrix(8, 3) <> 0 Then
         If .TextMatrix(8, 3) < 100 Or .TextMatrix(8, 3) > 100 Then
            .TextMatrix(8, 3) = "100" & "%"
         Else
            .TextMatrix(8, 3) = "100" & "%"
         End If
      End If

      If .TextMatrix(8, 5) <> 0 Then
         If .TextMatrix(8, 5) < 100 Or .TextMatrix(8, 5) > 100 Then
            .TextMatrix(8, 5) = "100" & "%"
         Else
            .TextMatrix(8, 5) = "100" & "%"
         End If
      End If
      .TextMatrix(8, 4) = Format(.TextMatrix(8, 4), "###,##0.00")

      'PERDIDA
      If .TextMatrix(12, 3) <> 0 Then
         If .TextMatrix(12, 3) < 100 Or .TextMatrix(12, 3) > 100 Then
            .TextMatrix(12, 3) = "100" & "%"
         Else
            .TextMatrix(12, 3) = "100" & "%"
         End If
      End If

      If .TextMatrix(12, 5) <> 0 Then
         If .TextMatrix(12, 5) < 100 Or .TextMatrix(12, 5) > 100 Then
            .TextMatrix(12, 5) = "100" & "%"
         Else
            .TextMatrix(12, 5) = "100" & "%"
         End If
      End If
      .TextMatrix(12, 4) = Format(.TextMatrix(12, 4), "###,##0.00")
   End With
  
   grd_LisDetCalif.Redraw = True
   grd_LisDetCalif.Enabled = True
   Call gs_UbicaGrid(grd_LisDetCalif, 2)
End Sub

Private Sub fs_Obtiene_Detalle_Producto()
   Dim r_int_Cant    As Integer
   Dim r_dbl_Monto   As Double
   
   grd_LisDetProd.Redraw = False
   Call gs_LimpiaGrid(grd_LisDetProd)
      
   With grd_LisDetProd
      .Rows = 31: .Cols = 10: .RowHeight(0) = 300 '27
      .FixedRows = 2
      .MergeCells = flexMergeFree
      
      .TextMatrix(0, 0) = "PRODUCTO"
      .TextMatrix(1, 0) = "PRODUCTO"
      .TextMatrix(0, 1) = "DEFICIENTE"
      .TextMatrix(0, 2) = "DEFICIENTE"
      .TextMatrix(0, 3) = "DEFICIENTE"
      .TextMatrix(1, 1) = "SITUACION"
      .TextMatrix(1, 2) = "N° CRED."
      .TextMatrix(1, 3) = "SALDO"
      
      .TextMatrix(0, 4) = "DUDOSO"
      .TextMatrix(0, 5) = "DUDOSO"
      .TextMatrix(0, 6) = "DUDOSO"
      .TextMatrix(1, 4) = "SITUACION"
      .TextMatrix(1, 5) = "N° CRED."
      .TextMatrix(1, 6) = "SALDO"
      
      .TextMatrix(0, 7) = "PERDIDA"
      .TextMatrix(0, 8) = "PERDIDA"
      .TextMatrix(0, 9) = "PERDIDA"
      .TextMatrix(1, 7) = "SITUACION"
      .TextMatrix(1, 8) = "N° CRED."
      .TextMatrix(1, 9) = "SALDO"
      
      .TextMatrix(2, 0) = "CRC-PBP"
      .TextMatrix(3, 0) = "CRC-PBP"
      .TextMatrix(4, 0) = "CRC-PBP"
      .TextMatrix(5, 0) = "CRC-PBP"
      .TextMatrix(6, 0) = "MICASITA"
      .TextMatrix(7, 0) = "MICASITA"
      .TextMatrix(8, 0) = "MICASITA"
      .TextMatrix(9, 0) = "MICASITA"
      .TextMatrix(10, 0) = "CME"
      .TextMatrix(11, 0) = "CME"
      .TextMatrix(12, 0) = "CME"
      .TextMatrix(13, 0) = "CME"
      .TextMatrix(14, 0) = "N.MIVIVIENDA"
      .TextMatrix(15, 0) = "N.MIVIVIENDA"
      .TextMatrix(16, 0) = "N.MIVIVIENDA"
      .TextMatrix(17, 0) = "N.MIVIVIENDA"
      .TextMatrix(18, 0) = "MICASA MAS"
      .TextMatrix(19, 0) = "MICASA MAS"
      .TextMatrix(20, 0) = "MICASA MAS"
      .TextMatrix(21, 0) = "MICASA MAS"
      .TextMatrix(22, 0) = "BBP"
      .TextMatrix(23, 0) = "BBP"
      .TextMatrix(24, 0) = "BBP"
      .TextMatrix(25, 0) = "BBP"
      
      .TextMatrix(26, 0) = "TECHO PROPIO"
      .TextMatrix(27, 0) = "TECHO PROPIO"
      .TextMatrix(28, 0) = "TECHO PROPIO"
      .TextMatrix(29, 0) = "TECHO PROPIO"
      
      
      .TextMatrix(30, 1) = "TOTAL" '26
      .TextMatrix(30, 4) = "TOTAL"
      .TextMatrix(30, 7) = "TOTAL"
      
      'DEFICIENTE
      .TextMatrix(2, 1) = "ALINEADO"
      .TextMatrix(3, 1) = "ALINEADO MOROSO"
      .TextMatrix(4, 1) = "MOROSO"
      .TextMatrix(5, 1) = "SUBTOTAL"

      .TextMatrix(6, 1) = "ALINEADO"
      .TextMatrix(7, 1) = "ALINEADO MOROSO"
      .TextMatrix(8, 1) = "MOROSO"
      .TextMatrix(9, 1) = "SUBTOTAL"

      .TextMatrix(10, 1) = "ALINEADO"
      .TextMatrix(11, 1) = "ALINEADO MOROSO"
      .TextMatrix(12, 1) = "MOROSO"
      .TextMatrix(13, 1) = "SUBTOTAL"

      .TextMatrix(14, 1) = "ALINEADO"
      .TextMatrix(15, 1) = "ALINEADO MOROSO"
      .TextMatrix(16, 1) = "MOROSO"
      .TextMatrix(17, 1) = "SUBTOTAL"

      .TextMatrix(18, 1) = "ALINEADO"
      .TextMatrix(19, 1) = "ALINEADO MOROSO"
      .TextMatrix(20, 1) = "MOROSO"
      .TextMatrix(21, 1) = "SUBTOTAL"

      .TextMatrix(22, 1) = "ALINEADO"
      .TextMatrix(23, 1) = "ALINEADO MOROSO"
      .TextMatrix(24, 1) = "MOROSO"
      .TextMatrix(25, 1) = "SUBTOTAL"
      
      .TextMatrix(26, 1) = "ALINEADO"
      .TextMatrix(27, 1) = "ALINEADO MOROSO"
      .TextMatrix(28, 1) = "MOROSO"
      .TextMatrix(29, 1) = "SUBTOTAL"

      'DUDOSO
      .TextMatrix(2, 4) = "ALINEADO"
      .TextMatrix(3, 4) = "ALINEADO MOROSO"
      .TextMatrix(4, 4) = "MOROSO"
      .TextMatrix(5, 4) = "SUBTOTAL"

      .TextMatrix(6, 4) = "ALINEADO"
      .TextMatrix(7, 4) = "ALINEADO MOROSO"
      .TextMatrix(8, 4) = "MOROSO"
      .TextMatrix(9, 4) = "SUBTOTAL"

      .TextMatrix(10, 4) = "ALINEADO"
      .TextMatrix(11, 4) = "ALINEADO MOROSO"
      .TextMatrix(12, 4) = "MOROSO"
      .TextMatrix(13, 4) = "SUBTOTAL"

      .TextMatrix(14, 4) = "ALINEADO"
      .TextMatrix(15, 4) = "ALINEADO MOROSO"
      .TextMatrix(16, 4) = "MOROSO"
      .TextMatrix(17, 4) = "SUBTOTAL"

      .TextMatrix(18, 4) = "ALINEADO"
      .TextMatrix(19, 4) = "ALINEADO MOROSO"
      .TextMatrix(20, 4) = "MOROSO"
      .TextMatrix(21, 4) = "SUBTOTAL"

      .TextMatrix(22, 4) = "ALINEADO"
      .TextMatrix(23, 4) = "ALINEADO MOROSO"
      .TextMatrix(24, 4) = "MOROSO"
      .TextMatrix(25, 4) = "SUBTOTAL"

      .TextMatrix(26, 4) = "ALINEADO"
      .TextMatrix(27, 4) = "ALINEADO MOROSO"
      .TextMatrix(28, 4) = "MOROSO"
      .TextMatrix(29, 4) = "SUBTOTAL"
      
      'PERDIDA
      .TextMatrix(2, 7) = "ALINEADO"
      .TextMatrix(3, 7) = "ALINEADO MOROSO"
      .TextMatrix(4, 7) = "MOROSO"
      .TextMatrix(5, 7) = "SUBTOTAL"

      .TextMatrix(6, 7) = "ALINEADO"
      .TextMatrix(7, 7) = "ALINEADO MOROSO"
      .TextMatrix(8, 7) = "MOROSO"
      .TextMatrix(9, 7) = "SUBTOTAL"

      .TextMatrix(10, 7) = "ALINEADO"
      .TextMatrix(11, 7) = "ALINEADO MOROSO"
      .TextMatrix(12, 7) = "MOROSO"
      .TextMatrix(13, 7) = "SUBTOTAL"

      .TextMatrix(14, 7) = "ALINEADO"
      .TextMatrix(15, 7) = "ALINEADO MOROSO"
      .TextMatrix(16, 7) = "MOROSO"
      .TextMatrix(17, 7) = "SUBTOTAL"

      .TextMatrix(18, 7) = "ALINEADO"
      .TextMatrix(19, 7) = "ALINEADO MOROSO"
      .TextMatrix(20, 7) = "MOROSO"
      .TextMatrix(21, 7) = "SUBTOTAL"

      .TextMatrix(22, 7) = "ALINEADO"
      .TextMatrix(23, 7) = "ALINEADO MOROSO"
      .TextMatrix(24, 7) = "MOROSO"
      .TextMatrix(25, 7) = "SUBTOTAL"

      .TextMatrix(26, 7) = "ALINEADO"
      .TextMatrix(27, 7) = "ALINEADO MOROSO"
      .TextMatrix(28, 7) = "MOROSO"
      .TextMatrix(29, 7) = "SUBTOTAL"
      
      .ColWidth(0) = 1900
      .ColWidth(1) = 1700
      .ColWidth(2) = 900
      .ColWidth(3) = 1200
      .ColWidth(4) = 1700
      .ColWidth(5) = 900
      .ColWidth(6) = 1200
      .ColWidth(7) = 1700
      .ColWidth(8) = 900
      .ColWidth(9) = 1200
      
      .FixedAlignment(0) = flexAlignCenterCenter
      .FixedAlignment(1) = flexAlignCenterCenter
      .FixedAlignment(2) = flexAlignCenterCenter
      .FixedAlignment(3) = flexAlignCenterCenter
      .FixedAlignment(4) = flexAlignCenterCenter
      .FixedAlignment(5) = flexAlignCenterCenter
      .FixedAlignment(6) = flexAlignCenterCenter
      .FixedAlignment(7) = flexAlignCenterCenter
      .FixedAlignment(8) = flexAlignCenterCenter
      .FixedAlignment(9) = flexAlignCenterCenter

      .MergeCol(0) = True
      .MergeRow(0) = True
      
      .Col = 1: .Row = 5: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 5: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 5: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 5: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 5: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 5: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 5: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 5: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 5: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      
      .Col = 1: .Row = 9: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 9: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 9: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 9: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 9: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 9: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 9: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 9: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 9: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      
      .Col = 1: .Row = 13: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 13: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 13: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 13: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 13: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 13: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 13: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 13: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 13: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      
      .Col = 1: .Row = 17: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 17: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 17: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 17: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 17: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 17: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 17: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 17: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 17: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      
      .Col = 1: .Row = 21: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 21: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 21: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 21: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 21: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 21: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 21: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 21: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 21: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      
      .Col = 1: .Row = 25: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 25: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 25: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 25: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 25: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 25: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 25: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 25: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 25: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      
      .Col = 1: .Row = 29: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 29: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 29: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 29: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 29: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 29: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 29: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 29: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 29: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      
      .Col = 1: .Row = 30: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 2: .Row = 30: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 3: .Row = 30: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 4: .Row = 30: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 5: .Row = 30: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 6: .Row = 30: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 7: .Row = 30: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 8: .Row = 30: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      .Col = 9: .Row = 30: .CellFontBold = True: .CellBackColor = RGB(238, 238, 238)
      
      '----------------------------------------------------------------------------
      'DEFICIENTE(CRC-PBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD = '001'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(2, 2) = Val(.TextMatrix(2, 2)) + g_rst_Princi!CONT
            .TextMatrix(2, 3) = Format(Val(Format(.TextMatrix(2, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .TextMatrix(3, 2) = Val(.TextMatrix(3, 2)) + g_rst_Princi!CONT
            .TextMatrix(3, 3) = Format(Val(Format(.TextMatrix(3, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .TextMatrix(4, 2) = Val(.TextMatrix(4, 2)) + g_rst_Princi!CONT
            .TextMatrix(4, 3) = Format(Val(Format(.TextMatrix(4, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DEFICIENTE(MICASITA)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD IN ('002','006','011')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(6, 2) = Val(.TextMatrix(6, 2)) + g_rst_Princi!CONT
            .TextMatrix(6, 3) = Format(Val(Format(.TextMatrix(6, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .TextMatrix(7, 2) = Val(.TextMatrix(7, 2)) + g_rst_Princi!CONT
            .TextMatrix(7, 3) = Format(Val(Format(.TextMatrix(7, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .TextMatrix(8, 2) = Val(.TextMatrix(8, 2)) + g_rst_Princi!CONT
            .TextMatrix(8, 3) = Format(Val(Format(.TextMatrix(8, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop

      'DEFICIENTE(CME)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD = '003'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(10, 2) = Val(.TextMatrix(10, 2)) + g_rst_Princi!CONT
            .TextMatrix(10, 3) = Format(Val(Format(.TextMatrix(10, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .TextMatrix(11, 2) = Val(.TextMatrix(11, 2)) + g_rst_Princi!CONT
            .TextMatrix(11, 3) = Format(Val(Format(.TextMatrix(11, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .TextMatrix(12, 2) = Val(.TextMatrix(12, 2)) + g_rst_Princi!CONT
            .TextMatrix(12, 3) = Format(Val(Format(.TextMatrix(12, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop

      'DEFICIENTE(N.MIVIVIENDA)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(14, 2) = Val(.TextMatrix(14, 2)) + g_rst_Princi!CONT
            .TextMatrix(14, 3) = Format(Val(Format(.TextMatrix(14, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .TextMatrix(15, 2) = Val(.TextMatrix(15, 2)) + g_rst_Princi!CONT
            .TextMatrix(15, 3) = Format(Val(Format(.TextMatrix(15, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .TextMatrix(16, 2) = Val(.TextMatrix(16, 2)) + g_rst_Princi!CONT
            .TextMatrix(16, 3) = Format(Val(Format(.TextMatrix(16, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop

      'DEFICIENTE(MICASAMAS)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD = '019'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(18, 2) = Val(.TextMatrix(18, 2)) + g_rst_Princi!CONT
            .TextMatrix(18, 3) = Format(Val(Format(.TextMatrix(18, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .TextMatrix(19, 2) = Val(.TextMatrix(19, 2)) + g_rst_Princi!CONT
            .TextMatrix(19, 3) = Format(Val(Format(.TextMatrix(19, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .TextMatrix(20, 2) = Val(.TextMatrix(20, 2)) + g_rst_Princi!CONT
            .TextMatrix(20, 3) = Format(Val(Format(.TextMatrix(20, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop

      'DEFICIENTE(BBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD IN ('021','022','023')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(22, 2) = Val(.TextMatrix(22, 2)) + g_rst_Princi!CONT
            .TextMatrix(22, 3) = Format(Val(Format(.TextMatrix(22, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .TextMatrix(23, 2) = Val(.TextMatrix(23, 2)) + g_rst_Princi!CONT
            .TextMatrix(23, 3) = Format(Val(Format(.TextMatrix(23, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .TextMatrix(24, 2) = Val(.TextMatrix(24, 2)) + g_rst_Princi!CONT
            .TextMatrix(24, 3) = Format(Val(Format(.TextMatrix(24, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop

      'DEFICIENTE(TECHO PROPIO)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 2 AND HIPCIE_CODPRD = '024' "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(26, 2) = Val(.TextMatrix(26, 2)) + g_rst_Princi!CONT
            .TextMatrix(26, 3) = Format(Val(Format(.TextMatrix(26, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Then
            .TextMatrix(27, 2) = Val(.TextMatrix(27, 2)) + g_rst_Princi!CONT
            .TextMatrix(27, 3) = Format(Val(Format(.TextMatrix(27, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 2 Then
            .TextMatrix(28, 2) = Val(.TextMatrix(28, 2)) + g_rst_Princi!CONT
            .TextMatrix(28, 3) = Format(Val(Format(.TextMatrix(28, 3), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop
      '----------------------------------------------------------------------------
      'DUDOSO(CRC-PBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD = '001'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(2, 5) = Val(.TextMatrix(2, 5)) + g_rst_Princi!CONT
            .TextMatrix(2, 6) = Format(Val(Format(.TextMatrix(2, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(3, 5) = Val(.TextMatrix(3, 5)) + g_rst_Princi!CONT
            .TextMatrix(3, 6) = Format(Val(Format(.TextMatrix(3, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .TextMatrix(4, 5) = Val(.TextMatrix(4, 5)) + g_rst_Princi!CONT
            .TextMatrix(4, 6) = Format(Val(Format(.TextMatrix(4, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DUDOSO(MICASITA)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD IN ('002','006','011')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(6, 5) = Val(.TextMatrix(6, 5)) + g_rst_Princi!CONT
            .TextMatrix(6, 6) = Format(Val(Format(.TextMatrix(6, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(7, 5) = Val(.TextMatrix(7, 5)) + g_rst_Princi!CONT
            .TextMatrix(7, 6) = Format(Val(Format(.TextMatrix(7, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .TextMatrix(8, 5) = Val(.TextMatrix(8, 5)) + g_rst_Princi!CONT
            .TextMatrix(8, 6) = Format(Val(Format(.TextMatrix(8, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DUDOSO(CME)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD = '003'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(10, 5) = Val(.TextMatrix(10, 5)) + g_rst_Princi!CONT
            .TextMatrix(10, 6) = Format(Val(Format(.TextMatrix(10, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(11, 5) = Val(.TextMatrix(11, 5)) + g_rst_Princi!CONT
            .TextMatrix(11, 6) = Format(Val(Format(.TextMatrix(11, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .TextMatrix(12, 5) = Val(.TextMatrix(12, 5)) + g_rst_Princi!CONT
            .TextMatrix(12, 6) = Format(Val(Format(.TextMatrix(12, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DUDOSO(N.MIVIVIENDA)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(14, 5) = Val(.TextMatrix(14, 5)) + g_rst_Princi!CONT
            .TextMatrix(14, 6) = Format(Val(Format(.TextMatrix(14, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(15, 5) = Val(.TextMatrix(15, 5)) + g_rst_Princi!CONT
            .TextMatrix(15, 6) = Format(Val(Format(.TextMatrix(15, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .TextMatrix(16, 5) = Val(.TextMatrix(16, 5)) + g_rst_Princi!CONT
            .TextMatrix(16, 6) = Format(Val(Format(.TextMatrix(16, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      'DUDOSO(MICASAMAS)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD = '019'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(18, 5) = Val(.TextMatrix(18, 5)) + g_rst_Princi!CONT
            .TextMatrix(18, 6) = Format(Val(Format(.TextMatrix(18, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(19, 5) = Val(.TextMatrix(19, 5)) + g_rst_Princi!CONT
            .TextMatrix(19, 6) = Format(Val(Format(.TextMatrix(19, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .TextMatrix(20, 5) = Val(.TextMatrix(20, 5)) + g_rst_Princi!CONT
            .TextMatrix(20, 6) = Format(Val(Format(.TextMatrix(20, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DUDOSO(BBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD IN ('021','022','023')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(22, 5) = Val(.TextMatrix(22, 5)) + g_rst_Princi!CONT
            .TextMatrix(22, 6) = Format(Val(Format(.TextMatrix(22, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(23, 5) = Val(.TextMatrix(23, 5)) + g_rst_Princi!CONT
            .TextMatrix(23, 6) = Format(Val(Format(.TextMatrix(23, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .TextMatrix(24, 5) = Val(.TextMatrix(24, 5)) + g_rst_Princi!CONT
            .TextMatrix(24, 6) = Format(Val(Format(.TextMatrix(24, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      'DUDOSO(TECHO PROPIO)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 3 AND HIPCIE_CODPRD = '024' "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(26, 5) = Val(.TextMatrix(26, 5)) + g_rst_Princi!CONT
            .TextMatrix(26, 6) = Format(Val(Format(.TextMatrix(26, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(27, 5) = Val(.TextMatrix(27, 5)) + g_rst_Princi!CONT
            .TextMatrix(27, 6) = Format(Val(Format(.TextMatrix(27, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 3 Then
            .TextMatrix(28, 5) = Val(.TextMatrix(28, 5)) + g_rst_Princi!CONT
            .TextMatrix(28, 6) = Format(Val(Format(.TextMatrix(28, 6), "#####0.00")) + g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop
      '----------------------------------------------------------------------------
      'PERDIDA(CRC-PBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD = '001'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"
 
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(2, 8) = g_rst_Princi!CONT
            .TextMatrix(2, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            .TextMatrix(3, 8) = r_int_Cant
            .TextMatrix(3, 9) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(4, 8) = g_rst_Princi!CONT
            .TextMatrix(4, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
      Loop

      r_int_Cant = 0
      r_dbl_Monto = 0
      
      'PERDIDA(MICASITA)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD IN ('002','006','011')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(6, 8) = g_rst_Princi!CONT
            .TextMatrix(6, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")

            .TextMatrix(7, 8) = r_int_Cant
            .TextMatrix(7, 9) = Format(r_dbl_Monto, "###,###,##0.00")
         End If

         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(8, 8) = g_rst_Princi!CONT
            .TextMatrix(8, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop

      r_int_Cant = 0
      r_dbl_Monto = 0
      'PERDIDA(CME)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD = '003'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(10, 8) = g_rst_Princi!CONT
            .TextMatrix(10, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")

            .TextMatrix(11, 8) = r_int_Cant
            .TextMatrix(11, 9) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(12, 8) = g_rst_Princi!CONT
            .TextMatrix(12, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop

      r_int_Cant = 0
      r_dbl_Monto = 0
      'PERDIDA(N.MIVIVIENDA)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(14, 8) = g_rst_Princi!CONT
            .TextMatrix(14, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")

            .TextMatrix(15, 8) = r_int_Cant
            .TextMatrix(15, 9) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(16, 8) = g_rst_Princi!CONT
            .TextMatrix(16, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop

      r_int_Cant = 0
      r_dbl_Monto = 0
      'PERDIDA(MICASAMAS)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD = '019'"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(18, 8) = g_rst_Princi!CONT
            .TextMatrix(18, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")

            .TextMatrix(19, 8) = r_int_Cant
            .TextMatrix(19, 9) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(20, 8) = g_rst_Princi!CONT
            .TextMatrix(20, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop

      r_int_Cant = 0
      r_dbl_Monto = 0
      'PERDIDA(BBP)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD IN ('021','022','023')"
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(22, 8) = g_rst_Princi!CONT
            .TextMatrix(22, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")

            .TextMatrix(23, 8) = r_int_Cant
            .TextMatrix(23, 9) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(24, 8) = g_rst_Princi!CONT
            .TextMatrix(24, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop
      
      r_int_Cant = 0
      r_dbl_Monto = 0
      'PERDIDA(TECHO PROPIO)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, HIPCIE_CLACLI, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
      g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
      g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " AND HIPCIE_CLAPRV = 4 AND HIPCIE_CODPRD = '024' "
      g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV, HIPCIE_CLACLI "
      g_str_Parame = g_str_Parame + " ORDER BY 1 DESC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPCIE_CLACLI = 0 Then
            .TextMatrix(26, 8) = g_rst_Princi!CONT
            .TextMatrix(26, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 1 Or g_rst_Princi!HIPCIE_CLACLI = 2 Or g_rst_Princi!HIPCIE_CLACLI = 3 Then
            r_int_Cant = r_int_Cant + g_rst_Princi!CONT
            r_dbl_Monto = r_dbl_Monto + Format(g_rst_Princi!SALDO, "###,###,##0.00")

            .TextMatrix(27, 8) = r_int_Cant
            .TextMatrix(27, 9) = Format(r_dbl_Monto, "###,###,##0.00")
         End If
         If g_rst_Princi!HIPCIE_CLACLI = 4 Then
            .TextMatrix(28, 8) = g_rst_Princi!CONT
            .TextMatrix(28, 9) = Format(g_rst_Princi!SALDO, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop
      
      
      'TOTAL DEFICIENTE
      .TextMatrix(5, 2) = Format(Val(.TextMatrix(2, 2)) + Val(.TextMatrix(3, 2)) + Val(.TextMatrix(4, 2)), "#,##0")
      .TextMatrix(5, 3) = Format(CDbl(IIf(.TextMatrix(2, 3) = "", 0, .TextMatrix(2, 3))) + CDbl(IIf(.TextMatrix(3, 3) = "", 0, .TextMatrix(3, 3))) + CDbl(IIf(.TextMatrix(4, 3) = "", 0, .TextMatrix(4, 3))), "###,###,##0.00")
      .TextMatrix(9, 2) = Format(Val(.TextMatrix(6, 2)) + Val(.TextMatrix(7, 2)) + Val(.TextMatrix(8, 2)), "#,##0")
      .TextMatrix(9, 3) = Format(CDbl(IIf(.TextMatrix(6, 3) = "", 0, .TextMatrix(6, 3))) + CDbl(IIf(.TextMatrix(7, 3) = "", 0, .TextMatrix(7, 3))) + CDbl(IIf(.TextMatrix(8, 3) = "", 0, .TextMatrix(8, 3))), "###,###,##0.00")
      .TextMatrix(13, 2) = Format(Val(.TextMatrix(10, 2)) + Val(.TextMatrix(11, 2)) + Val(.TextMatrix(12, 2)), "#,##0")
      .TextMatrix(13, 3) = Format(CDbl(IIf(.TextMatrix(10, 3) = "", 0, .TextMatrix(10, 3))) + CDbl(IIf(.TextMatrix(11, 3) = "", 0, .TextMatrix(11, 3))) + CDbl(IIf(.TextMatrix(12, 3) = "", 0, .TextMatrix(12, 3))), "###,###,##0.00")
      .TextMatrix(17, 2) = Format(Val(.TextMatrix(14, 2)) + Val(.TextMatrix(15, 2)) + Val(.TextMatrix(16, 2)), "#,##0")
      .TextMatrix(17, 3) = Format(CDbl(IIf(.TextMatrix(14, 3) = "", 0, .TextMatrix(14, 3))) + CDbl(IIf(.TextMatrix(15, 3) = "", 0, .TextMatrix(15, 3))) + CDbl(IIf(.TextMatrix(16, 3) = "", 0, .TextMatrix(16, 3))), "###,###,##0.00")
      .TextMatrix(21, 2) = Format(Val(.TextMatrix(18, 2)) + Val(.TextMatrix(19, 2)) + Val(.TextMatrix(20, 2)), "#,##0")
      .TextMatrix(21, 3) = Format(CDbl(IIf(.TextMatrix(18, 3) = "", 0, .TextMatrix(18, 3))) + CDbl(IIf(.TextMatrix(19, 3) = "", 0, .TextMatrix(19, 3))) + CDbl(IIf(.TextMatrix(20, 3) = "", 0, .TextMatrix(20, 3))), "###,###,##0.00")
      .TextMatrix(25, 2) = Format(Val(.TextMatrix(22, 2)) + Val(.TextMatrix(23, 2)) + Val(.TextMatrix(24, 2)), "#,##0")
      .TextMatrix(25, 3) = Format(CDbl(IIf(.TextMatrix(22, 3) = "", 0, .TextMatrix(22, 3))) + CDbl(IIf(.TextMatrix(23, 3) = "", 0, .TextMatrix(23, 3))) + CDbl(IIf(.TextMatrix(24, 3) = "", 0, .TextMatrix(24, 3))), "###,###,##0.00")
      .TextMatrix(29, 2) = Format(Val(.TextMatrix(26, 2)) + Val(.TextMatrix(27, 2)) + Val(.TextMatrix(28, 2)), "#,##0")
      .TextMatrix(29, 3) = Format(CDbl(IIf(.TextMatrix(26, 3) = "", 0, .TextMatrix(26, 3))) + CDbl(IIf(.TextMatrix(27, 3) = "", 0, .TextMatrix(27, 3))) + CDbl(IIf(.TextMatrix(28, 3) = "", 0, .TextMatrix(28, 3))), "###,###,##0.00")

      'TOTAL DUDOSO
      .TextMatrix(5, 5) = Format(Val(.TextMatrix(2, 5)) + Val(.TextMatrix(3, 5)) + Val(.TextMatrix(4, 5)), "#,##0")
      .TextMatrix(5, 6) = Format(CDbl(IIf(.TextMatrix(2, 6) = "", 0, .TextMatrix(2, 6))) + CDbl(IIf(.TextMatrix(3, 6) = "", 0, .TextMatrix(3, 6))) + CDbl(IIf(.TextMatrix(4, 6) = "", 0, .TextMatrix(4, 6))), "###,###,##0.00")
      .TextMatrix(9, 5) = Format(Val(.TextMatrix(6, 5)) + Val(.TextMatrix(7, 5)) + Val(.TextMatrix(8, 5)), "#,##0")
      .TextMatrix(9, 6) = Format(CDbl(IIf(.TextMatrix(6, 6) = "", 0, .TextMatrix(6, 6))) + CDbl(IIf(.TextMatrix(7, 6) = "", 0, .TextMatrix(7, 6))) + CDbl(IIf(.TextMatrix(8, 6) = "", 0, .TextMatrix(8, 6))), "###,###,##0.00")
      .TextMatrix(13, 5) = Format(Val(.TextMatrix(10, 5)) + Val(.TextMatrix(11, 5)) + Val(.TextMatrix(12, 5)), "#,##0")
      .TextMatrix(13, 6) = Format(CDbl(IIf(.TextMatrix(10, 6) = "", 0, .TextMatrix(10, 6))) + CDbl(IIf(.TextMatrix(11, 6) = "", 0, .TextMatrix(11, 6))) + CDbl(IIf(.TextMatrix(12, 6) = "", 0, .TextMatrix(12, 6))), "###,###,##0.00")
      .TextMatrix(17, 5) = Format(Val(.TextMatrix(14, 5)) + Val(.TextMatrix(15, 5)) + Val(.TextMatrix(16, 5)), "#,##0")
      .TextMatrix(17, 6) = Format(CDbl(IIf(.TextMatrix(14, 6) = "", 0, .TextMatrix(14, 6))) + CDbl(IIf(.TextMatrix(15, 6) = "", 0, .TextMatrix(15, 6))) + CDbl(IIf(.TextMatrix(16, 6) = "", 0, .TextMatrix(16, 6))), "###,###,##0.00")
      .TextMatrix(21, 5) = Format(Val(.TextMatrix(18, 5)) + Val(.TextMatrix(19, 5)) + Val(.TextMatrix(20, 5)), "#,##0")
      .TextMatrix(21, 6) = Format(CDbl(IIf(.TextMatrix(18, 6) = "", 0, .TextMatrix(18, 6))) + CDbl(IIf(.TextMatrix(19, 6) = "", 0, .TextMatrix(19, 6))) + CDbl(IIf(.TextMatrix(20, 6) = "", 0, .TextMatrix(20, 6))), "###,###,##0.00")
      .TextMatrix(25, 5) = Format(Val(.TextMatrix(22, 5)) + Val(.TextMatrix(23, 5)) + Val(.TextMatrix(24, 5)), "#,##0")
      .TextMatrix(25, 6) = Format(CDbl(IIf(.TextMatrix(22, 6) = "", 0, .TextMatrix(22, 6))) + CDbl(IIf(.TextMatrix(23, 6) = "", 0, .TextMatrix(23, 6))) + CDbl(IIf(.TextMatrix(24, 6) = "", 0, .TextMatrix(24, 6))), "###,###,##0.00")

      .TextMatrix(29, 5) = Format(Val(.TextMatrix(26, 5)) + Val(.TextMatrix(27, 5)) + Val(.TextMatrix(28, 5)), "#,##0")
      .TextMatrix(29, 6) = Format(CDbl(IIf(.TextMatrix(26, 6) = "", 0, .TextMatrix(26, 6))) + CDbl(IIf(.TextMatrix(27, 6) = "", 0, .TextMatrix(27, 6))) + CDbl(IIf(.TextMatrix(28, 6) = "", 0, .TextMatrix(28, 6))), "###,###,##0.00")

      'TOTAL PERDIDA
      .TextMatrix(5, 8) = Format(Val(.TextMatrix(2, 8)) + Val(.TextMatrix(3, 8)) + Val(.TextMatrix(4, 8)), "#,##0")
      .TextMatrix(5, 9) = Format(CDbl(IIf(.TextMatrix(2, 9) = "", 0, .TextMatrix(2, 9))) + CDbl(IIf(.TextMatrix(3, 9) = "", 0, .TextMatrix(3, 9))) + CDbl(IIf(.TextMatrix(4, 9) = "", 0, .TextMatrix(4, 9))), "###,###,##0.00")
      .TextMatrix(9, 8) = Format(Val(.TextMatrix(6, 8)) + Val(.TextMatrix(7, 8)) + Val(.TextMatrix(8, 8)), "#,##0")
      .TextMatrix(9, 9) = Format(CDbl(IIf(.TextMatrix(6, 9) = "", 0, .TextMatrix(6, 9))) + CDbl(IIf(.TextMatrix(7, 9) = "", 0, .TextMatrix(7, 9))) + CDbl(IIf(.TextMatrix(8, 9) = "", 0, .TextMatrix(8, 9))), "###,###,##0.00")
      .TextMatrix(13, 8) = Format(Val(.TextMatrix(10, 8)) + Val(.TextMatrix(11, 8)) + Val(.TextMatrix(12, 8)), "#,##0")
      .TextMatrix(13, 9) = Format(CDbl(IIf(.TextMatrix(10, 9) = "", 0, .TextMatrix(10, 9))) + CDbl(IIf(.TextMatrix(11, 9) = "", 0, .TextMatrix(11, 9))) + CDbl(IIf(.TextMatrix(12, 9) = "", 0, .TextMatrix(12, 9))), "###,###,##0.00")
      .TextMatrix(17, 8) = Format(Val(.TextMatrix(14, 8)) + Val(.TextMatrix(15, 8)) + Val(.TextMatrix(16, 8)), "#,##0")
      .TextMatrix(17, 9) = Format(CDbl(IIf(.TextMatrix(14, 9) = "", 0, .TextMatrix(14, 9))) + CDbl(IIf(.TextMatrix(15, 9) = "", 0, .TextMatrix(15, 9))) + CDbl(IIf(.TextMatrix(16, 9) = "", 0, .TextMatrix(16, 9))), "###,###,##0.00")
      .TextMatrix(21, 8) = Format(Val(.TextMatrix(18, 8)) + Val(.TextMatrix(19, 8)) + Val(.TextMatrix(20, 8)), "#,##0")
      .TextMatrix(21, 9) = Format(CDbl(IIf(.TextMatrix(18, 9) = "", 0, .TextMatrix(18, 9))) + CDbl(IIf(.TextMatrix(19, 9) = "", 0, .TextMatrix(19, 9))) + CDbl(IIf(.TextMatrix(20, 9) = "", 0, .TextMatrix(20, 9))), "###,###,##0.00")
      .TextMatrix(25, 8) = Format(Val(.TextMatrix(22, 8)) + Val(.TextMatrix(23, 8)) + Val(.TextMatrix(24, 8)), "#,##0")
      .TextMatrix(25, 9) = Format(CDbl(IIf(.TextMatrix(22, 9) = "", 0, .TextMatrix(22, 9))) + CDbl(IIf(.TextMatrix(23, 9) = "", 0, .TextMatrix(23, 9))) + CDbl(IIf(.TextMatrix(24, 9) = "", 0, .TextMatrix(24, 9))), "###,###,##0.00")

      .TextMatrix(29, 8) = Format(Val(.TextMatrix(26, 8)) + Val(.TextMatrix(27, 8)) + Val(.TextMatrix(28, 8)), "#,##0")
      .TextMatrix(29, 9) = Format(CDbl(IIf(.TextMatrix(26, 9) = "", 0, .TextMatrix(26, 9))) + CDbl(IIf(.TextMatrix(27, 9) = "", 0, .TextMatrix(27, 9))) + CDbl(IIf(.TextMatrix(28, 9) = "", 0, .TextMatrix(28, 9))), "###,###,##0.00")

      'TOTAL FINAL
      .TextMatrix(30, 2) = Format(Val(.TextMatrix(5, 2)) + Val(.TextMatrix(9, 2)) + Val(.TextMatrix(13, 2)) + Val(.TextMatrix(17, 2)) + Val(.TextMatrix(21, 2)) + Val(.TextMatrix(25, 2)) + Val(.TextMatrix(29, 2)), "#,##0")
      .TextMatrix(30, 3) = Format(CDbl(.TextMatrix(5, 3)) + CDbl(.TextMatrix(9, 3)) + CDbl(.TextMatrix(13, 3)) + CDbl(.TextMatrix(17, 3)) + CDbl(.TextMatrix(21, 3)) + CDbl(.TextMatrix(25, 3)) + CDbl(.TextMatrix(29, 3)), "###,###,##0.00")
      .TextMatrix(30, 5) = Format(Val(.TextMatrix(5, 5)) + Val(.TextMatrix(9, 5)) + Val(.TextMatrix(13, 5)) + Val(.TextMatrix(17, 5)) + Val(.TextMatrix(21, 5)) + Val(.TextMatrix(25, 5)) + Val(.TextMatrix(29, 5)), "#,##0")
      .TextMatrix(30, 6) = Format(CDbl(.TextMatrix(5, 6)) + CDbl(.TextMatrix(9, 6)) + CDbl(.TextMatrix(13, 6)) + CDbl(.TextMatrix(17, 6)) + CDbl(.TextMatrix(21, 6)) + CDbl(.TextMatrix(25, 6)) + CDbl(.TextMatrix(29, 6)), "###,###,##0.00")
      .TextMatrix(30, 8) = Format(Val(.TextMatrix(5, 8)) + Val(.TextMatrix(9, 8)) + Val(.TextMatrix(13, 8)) + Val(.TextMatrix(17, 8)) + Val(.TextMatrix(21, 8)) + Val(.TextMatrix(25, 8)) + Val(.TextMatrix(29, 8)), "#,##0")
      .TextMatrix(30, 9) = Format(CDbl(.TextMatrix(5, 9)) + CDbl(.TextMatrix(9, 9)) + CDbl(.TextMatrix(13, 9)) + CDbl(.TextMatrix(17, 9)) + CDbl(.TextMatrix(21, 9)) + CDbl(.TextMatrix(25, 9)) + CDbl(.TextMatrix(29, 9)), "###,###,##0.00")
   End With
  
   grd_LisDetProd.Redraw = True
   grd_LisDetProd.Enabled = True
   Call gs_UbicaGrid(grd_LisDetProd, 2)
End Sub

Private Sub fs_GenExc()
Dim r_int_NroFil     As Integer
Dim r_int_NoFlLi     As Integer
Dim r_int_TotReg     As Integer

   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 3
   r_obj_Excel.Workbooks.Add
   
   Call fs_GenExc_Clasificacion
   Call fs_GenExc_Detalle_Clasificacion
   Call fs_GenExc_Detalle_Producto
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   
   Call fs_Obtiene_Clasificacion
   Call fs_Obtiene_Detalle_Clasificacion
   Call fs_Obtiene_Detalle_Producto
   
   Call fs_Activa(True)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = "" Then
      MsgBox "Debe seleccionar el Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(False)
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If CInt(ipp_PerAno.Text) >= 2007 Then
         Call gs_SetFocus(cmd_Buscar)
      End If
   End If
End Sub

