VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Ges_TecPro_14 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12795
   Icon            =   "OpeTra_frm_836.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6165
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   12795
      _Version        =   65536
      _ExtentX        =   22569
      _ExtentY        =   10874
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   6075
         Left            =   0
         TabIndex        =   1
         Top             =   30
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
         _ExtentY        =   10716
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   675
            Left            =   30
            TabIndex        =   2
            Top             =   750
            Width           =   12675
            _Version        =   65536
            _ExtentX        =   22357
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
            Begin VB.CommandButton cmd_Salida 
               Height          =   585
               Left            =   12060
               Picture         =   "OpeTra_frm_836.frx":000C
               Style           =   1  'Graphical
               TabIndex        =   28
               ToolTipText     =   "Salir"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Acepta 
               Height          =   585
               Left            =   11490
               Picture         =   "OpeTra_frm_836.frx":044E
               Style           =   1  'Graphical
               TabIndex        =   27
               ToolTipText     =   "Aceptar Datos"
               Top             =   30
               Width           =   585
            End
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   3345
            Left            =   30
            TabIndex        =   3
            Top             =   2640
            Width           =   12675
            _Version        =   65536
            _ExtentX        =   22357
            _ExtentY        =   5900
            _StockProps     =   15
            Caption         =   "1395"
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
               Height          =   2955
               Left            =   60
               TabIndex        =   16
               Top             =   360
               Width           =   12570
               _ExtentX        =   22172
               _ExtentY        =   5212
               _Version        =   393216
               Rows            =   45
               Cols            =   9
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_Item 
               Height          =   315
               Left            =   60
               TabIndex        =   17
               Top             =   60
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Item"
               ForeColor       =   16777215
               BackColor       =   16384
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
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Seleccionar 
               Height          =   315
               Left            =   10890
               TabIndex        =   18
               Top             =   60
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "  Seleccionar"
               ForeColor       =   16777215
               BackColor       =   16384
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
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   1
               Begin VB.CheckBox chkSeleccionar 
                  BackColor       =   &H00004000&
                  Caption         =   "Check1"
                  Height          =   255
                  Left            =   1110
                  TabIndex        =   19
                  Top             =   0
                  Width           =   255
               End
            End
            Begin Threed.SSPanel pnl_FecEmi 
               Height          =   315
               Left            =   2160
               TabIndex        =   20
               Top             =   60
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Fecha Emisión"
               ForeColor       =   16777215
               BackColor       =   16384
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
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_FecVto 
               Height          =   315
               Left            =   4590
               TabIndex        =   21
               Top             =   60
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Fecha Vcto."
               ForeColor       =   16777215
               BackColor       =   16384
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
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_NroCar 
               Height          =   315
               Left            =   660
               TabIndex        =   22
               Top             =   60
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Número"
               ForeColor       =   16777215
               BackColor       =   16384
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
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Plazo 
               Height          =   315
               Left            =   3660
               TabIndex        =   23
               Top             =   60
               Width           =   945
               _Version        =   65536
               _ExtentX        =   1667
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Plazo"
               ForeColor       =   16777215
               BackColor       =   16384
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
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Moneda 
               Height          =   315
               Left            =   6090
               TabIndex        =   24
               Top             =   60
               Width           =   2085
               _Version        =   65536
               _ExtentX        =   3678
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Moneda"
               ForeColor       =   16777215
               BackColor       =   16384
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
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_ImpCFi 
               Height          =   315
               Left            =   8160
               TabIndex        =   25
               Top             =   60
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Importe"
               ForeColor       =   16777215
               BackColor       =   16384
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
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_ImpGar 
               Height          =   315
               Left            =   9540
               TabIndex        =   26
               Top             =   60
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Garantizado"
               ForeColor       =   16777215
               BackColor       =   16384
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
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin VB.Line Line1 
               X1              =   3900
               X2              =   3930
               Y1              =   4710
               Y2              =   4740
            End
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   675
            Left            =   30
            TabIndex        =   4
            Top             =   30
            Width           =   12675
            _Version        =   65536
            _ExtentX        =   22357
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
               Left            =   630
               TabIndex        =   5
               Top             =   30
               Width           =   4305
               _Version        =   65536
               _ExtentX        =   7594
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Gestión de Crédito Hipotecario"
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
            Begin Threed.SSPanel pnl_Descri 
               Height          =   315
               Left            =   630
               TabIndex        =   6
               Top             =   330
               Width           =   4575
               _Version        =   65536
               _ExtentX        =   8070
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Techo Propio - Asociar Cartas Fianza y Adendas"
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
               Picture         =   "OpeTra_frm_836.frx":0758
               Top             =   60
               Width           =   480
            End
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   1125
            Left            =   30
            TabIndex        =   7
            Top             =   1470
            Width           =   12675
            _Version        =   65536
            _ExtentX        =   22357
            _ExtentY        =   1984
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
               Left            =   1740
               TabIndex        =   8
               Top             =   450
               Width           =   5625
               _Version        =   65536
               _ExtentX        =   9922
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_TipDoc 
               Height          =   315
               Left            =   1740
               TabIndex        =   9
               Top             =   120
               Width           =   5625
               _Version        =   65536
               _ExtentX        =   9922
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_NroDoc 
               Height          =   315
               Left            =   9390
               TabIndex        =   10
               Top             =   120
               Width           =   2895
               _Version        =   65536
               _ExtentX        =   5106
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_TipEmp 
               Height          =   315
               Left            =   1740
               TabIndex        =   11
               Top             =   780
               Width           =   2895
               _Version        =   65536
               _ExtentX        =   5106
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin VB.Label lbl_TipEmp 
               Caption         =   "Tipo Empresa:"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   780
               Width           =   1335
            End
            Begin VB.Label lbl_RazSoc 
               Caption         =   "Razón Social:"
               Height          =   255
               Left            =   150
               TabIndex        =   14
               Top             =   435
               Width           =   1035
            End
            Begin VB.Label lbl_NumDoc 
               Caption         =   "Nro. Documento:"
               Height          =   225
               Left            =   7770
               TabIndex        =   13
               Top             =   135
               Width           =   1335
            End
            Begin VB.Label lbl_TipDoc 
               Caption         =   "Tipo Documento:"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Type r_Arr_CarFia
'   CarFia_NumCFi        As String
'   CarFia_FecEmi        As String
'   CarFia_PlaCFi        As String
'   CarFia_FecVct        As String
'   CarFia_Situac        As Integer
'   CarFia_Moneda        As String
'   CarFia_ImpCFi        As Double
'   CarFia_ImpGar        As Double
'End Type
'Dim arr_CarFia()        As modprc_g_tpo_Genera

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 8) = ""
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 8) = "X"
         Next r_Fila
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub cmd_Grabar_Click()

End Sub

Private Sub cmd_Acepta_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer
Dim r_str_NumCFi        As String

   'valida selección
   r_int_ConSel = 0
   r_str_NumCFi = ""
   ReDim modatecli_g_arr_DocCre(0)
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 8) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionado Cartas Fianza o Adendas a Asociar.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirma
   If MsgBox("¿Está seguro de Asociar las Cartas Fianza o Adendas seleccionadas?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 8) = "X" Then
         r_str_NumCFi = grd_Listad.TextMatrix(r_int_Contad, 1) & "|" & r_str_NumCFi
      End If
   Next
   
   ReDim Preserve modatecli_g_arr_DocCre(1)
   r_str_NumCFi = Mid(r_str_NumCFi, 1, InStrRev(r_str_NumCFi, "|") - 1)
   modatecli_g_arr_DocCre(1).DocCre_CodIte = r_str_NumCFi

   chkSeleccionar.Value = False
   frm_Ges_TecPro_05.pnl_NumCFi.Caption = frm_Ges_TecPro_05.pnl_NumCFi.Caption & "|" & modatecli_g_arr_DocCre(1).DocCre_CodIte
   Call fs_Limpia
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   Call gs_CentraForm(Me)
  
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 575  'ITEM
   grd_Listad.ColWidth(1) = 1510 'CARTA FIANZA
   grd_Listad.ColWidth(2) = 1490 'FECHA DE EMISION
   grd_Listad.ColWidth(3) = 945 'PLAZO
   grd_Listad.ColWidth(4) = 1490 'FECHA DE VENCIMIENTO
   grd_Listad.ColWidth(5) = 2070 'MONEDA
   grd_Listad.ColWidth(6) = 1395 'IMPORTE
   grd_Listad.ColWidth(7) = 1340 'GARANTIZADO
   grd_Listad.ColWidth(8) = 1430 'SELECCIONAR
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   
   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", moddat_g_int_TipDoc)
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri

End Sub
Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub
Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      If grd_Listad.TextMatrix(grd_Listad.Row, 8) = "X" Then
         grd_Listad.TextMatrix(grd_Listad.Row, 8) = ""
      Else
         grd_Listad.TextMatrix(grd_Listad.Row, 8) = "X"
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub
Private Sub fs_Buscar(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
Dim r_str_Cadena  As String
Dim r_str_CadAux  As String
Dim r_str_CadRef  As String

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT MAEGAR_NUMREF AS NUMREF FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "    WHERE MAEGAR_TIPDOC = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "      AND MAEGAR_NUMDOC = '" & CStr(p_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      GoTo Ingresar 'Exit Sub
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      r_str_Cadena = Trim(g_rst_GenAux!NUMREF)
      
      'Obtiene los números de referencias actuales
      r_str_CadAux = Trim(r_str_Cadena)
      If Len(r_str_CadAux) > 1 Then
         While InStr(r_str_CadAux, "|")
            r_str_CadRef = Trim(Mid(r_str_CadAux, 1, InStr(r_str_CadAux, "|") - 1))
            r_str_CadRef = fs_Obtener_NumRef(r_str_CadRef)
            r_str_CadAux = Trim(Mid(r_str_CadAux, InStr(r_str_CadAux, "|") + 1))
         Wend
         
         r_str_CadRef = r_str_CadRef & "|" & fs_Obtener_NumRef(r_str_CadAux)
         If InStr(r_str_CadRef, "|") = 1 Then
            r_str_CadRef = Replace(r_str_CadRef, "|", "")
         End If
         r_str_CadRef = Replace(r_str_CadRef, "|", "' , '")
      End If
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing

Ingresar:

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAECFI_NUMREF, MAECFI_EMIFIA, MAECFI_PLZFIA, MAECFI_VTOFIA, MAECFI_MONFIA, MAECFI_IMPFIA, MAECFI_GARFIA, MAECFI_NUMANT "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "  WHERE MAECFI_TIPDOC = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND MAECFI_NUMDOC = '" & CStr(p_NumDoc) & "' "
   If r_str_Cadena <> "" Then
      g_str_Parame = g_str_Parame & "    AND MAECFI_NUMREF NOT IN ('" & r_str_CadRef & "')" 'r_str_Cadena
   End If
   g_str_Parame = g_str_Parame & "    AND MAECFI_SITUAC = 1 "
            
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
      grd_Listad.Redraw = True
     Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
          
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = grd_Listad.Row + 1
         
         grd_Listad.Col = 1
         If Not IsNull(g_rst_Princi!MAECFI_NUMANT) Then
            grd_Listad.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!MAECFI_NUMANT)), Mid(g_rst_Princi!MAECFI_NUMANT, 1, 1))
         Else
            grd_Listad.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!MAECFI_NUMREF)), Mid(g_rst_Princi!MAECFI_NUMREF, 1, 1))
         End If
         
         grd_Listad.Col = 2
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_EMIFIA)), "dd/mm/yyyy")
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!MAECFI_PLZFIA)
         
         grd_Listad.Col = 4
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
         
         grd_Listad.Col = 5
         grd_Listad.Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!MAECFI_MONFIA)
         
         grd_Listad.Col = 6
         grd_Listad.Text = Format(CStr(g_rst_Princi!MAECFI_IMPFIA), "###,###,###,##0.00")
         
         grd_Listad.Col = 7
         grd_Listad.Text = Format(CStr(g_rst_Princi!MAECFI_GARFIA), "###,###,###,##0.00")
                         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)

End Sub
'Private Function fs_Formato_NumRef(ByVal p_Numref As String) As String
'   p_Numref = Format(p_Numref, "0000000000")
'   'fs_Formato_NumRef = Left(p_Numref, 4) & "-" & Mid(p_Numref, 5, 2) & "-" & Right(p_Numref, 4)
'   fs_Formato_NumRef = Mid(p_Numref, 1, 1) & Mid(p_Numref, 2, 2) & "-" & Mid(p_Numref, 4, 2) & "-" & Right(p_Numref, 5)
'End Function
Private Function fs_Obtener_NumRef(ByVal p_NumRef As String) As String ', ByVal p_Tipo As Integer
   fs_Obtener_NumRef = ""
   p_NumRef = Format(p_NumRef, "0000000000")
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAECFI_NUMREF, MAECFI_NUMANT "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "  WHERE MAECFI_SITUAC = 1 "
   
   'If p_Tipo = 0 Then
      g_str_Parame = g_str_Parame & "     AND MAECFI_NUMANT = '" & p_NumRef & "'"
   'End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     fs_Obtener_NumRef = p_NumRef
     Exit Function
   End If
   
   fs_Obtener_NumRef = g_rst_Princi!MAECFI_NUMREF
End Function
