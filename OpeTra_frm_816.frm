VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Con_PreSeg_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16725
   Icon            =   "OpeTra_frm_816.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   16725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16740
      _Version        =   65536
      _ExtentX        =   29527
      _ExtentY        =   16880
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   1
         Top             =   810
         Width           =   16635
         _Version        =   65536
         _ExtentX        =   29342
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
            Left            =   2430
            Picture         =   "OpeTra_frm_816.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Exportar datos a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_RepCro 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_816.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Recepcionar Cronogramas"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_AboPag 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_816.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Abonar a COFIDE"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_816.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Imprimir Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   16035
            Picture         =   "OpeTra_frm_816.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_816.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   16635
         _Version        =   65536
         _ExtentX        =   29342
         _ExtentY        =   1244
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
            Left            =   660
            TabIndex        =   5
            Top             =   60
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Operaciones Financieras"
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
            Height          =   315
            Left            =   660
            TabIndex        =   6
            Top             =   360
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Seguimiento de Prepagos"
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
            Left            =   10920
            Top             =   120
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
            Left            =   60
            Picture         =   "OpeTra_frm_816.frx":14B8
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6165
         Left            =   60
         TabIndex        =   7
         Top             =   1515
         Width           =   16635
         _Version        =   65536
         _ExtentX        =   29342
         _ExtentY        =   10874
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
            Height          =   5730
            Left            =   45
            TabIndex        =   8
            Top             =   405
            Width           =   16560
            _ExtentX        =   29210
            _ExtentY        =   10107
            _Version        =   393216
            Rows            =   26
            Cols            =   18
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   90
            TabIndex        =   9
            Top             =   90
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
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
         Begin Threed.SSPanel pnl_Tit_TipPpg 
            Height          =   285
            Left            =   6495
            TabIndex        =   10
            Top             =   90
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Prepago"
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
         Begin Threed.SSPanel pnl_Tit_Import 
            Height          =   285
            Left            =   10980
            TabIndex        =   11
            Top             =   90
            Width           =   1620
            _Version        =   65536
            _ExtentX        =   2857
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
         Begin Threed.SSPanel pnl_Tit_FecPro 
            Height          =   285
            Left            =   9825
            TabIndex        =   12
            Top             =   90
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Proceso"
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
         Begin Threed.SSPanel pnl_Tit_DoiCli 
            Height          =   285
            Left            =   1200
            TabIndex        =   13
            Top             =   90
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
            Left            =   2415
            TabIndex        =   14
            Top             =   90
            Width           =   4200
            _Version        =   65536
            _ExtentX        =   7408
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Cliente"
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
         Begin Threed.SSPanel pnl_Tit_FecPpg 
            Height          =   285
            Left            =   8670
            TabIndex        =   15
            Top             =   90
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Prepago"
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
         Begin Threed.SSPanel SSPanel81 
            Height          =   285
            Left            =   14940
            TabIndex        =   17
            Top             =   90
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2469
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "  Seleccionar"
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
            Alignment       =   1
            Begin VB.CheckBox chkSeleccionar 
               BackColor       =   &H00004000&
               Caption         =   "Check1"
               Height          =   255
               Left            =   1110
               TabIndex        =   18
               Top             =   0
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_Tit_Estado 
            Height          =   285
            Left            =   12570
            TabIndex        =   19
            Top             =   90
            Width           =   2400
            _Version        =   65536
            _ExtentX        =   4233
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Estado"
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   900
         Left            =   60
         TabIndex        =   22
         Top             =   7710
         Width           =   16605
         _Version        =   65536
         _ExtentX        =   29289
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   315
            Left            =   240
            TabIndex        =   23
            Top             =   120
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "PREPAGO GENERADO"
            ForeColor       =   255
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   315
            Left            =   240
            TabIndex        =   24
            Top             =   480
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ABONO COFIDE"
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
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_AboCOF 
            Height          =   315
            Left            =   2820
            TabIndex        =   25
            Top             =   480
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_PPGRea 
            Height          =   315
            Left            =   2820
            TabIndex        =   26
            Top             =   120
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_RecCro 
            Height          =   315
            Left            =   9930
            TabIndex        =   27
            Top             =   480
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CarCOF 
            Height          =   315
            Left            =   9930
            TabIndex        =   28
            Top             =   120
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   0
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
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   315
            Left            =   6000
            TabIndex        =   29
            Top             =   480
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "RECEPCION DE CALENDARIO"
            ForeColor       =   33023
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   315
            Left            =   6000
            TabIndex        =   30
            Top             =   120
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "ENVIADO A COFIDE (CARTA GENERADA)"
            ForeColor       =   16711680
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
            Font3D          =   2
            Alignment       =   1
         End
      End
   End
End
Attribute VB_Name = "frm_Con_PreSeg_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private l_rst_Prepagos     As ADODB.Recordset

Private Sub cmd_AboPag_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer

   'valida selección
   r_int_ConSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 14) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionado Solicitudes para abonar.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Valida que las solicitudes seleccionadas solo sean --> Debe tener estado de Enviado a COFIDE (Est = 2) para que pase a Abono a COFIDE
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If (grd_Listad.TextMatrix(r_int_Contad, 14) = "X") Then
         If Not fs_Valida_EstPrepago(grd_Listad.TextMatrix(r_int_Contad, 13), 2) Then
            If Trim(grd_Listad.TextMatrix(r_int_Contad, 13)) <> "ABONO A COFIDE" Then
               MsgBox "La solicitud " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede Abonar, porque no se encuentra en instancia ENVIO A COFIDE.", vbInformation, modgen_g_str_NomPlt
            Else
               MsgBox "La solicitud " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede Abonar, porque ya se encuentra en instancia ABONO A COFIDE.", vbInformation, modgen_g_str_NomPlt
            End If
            Screen.MousePointer = 0
            Exit Sub
         End If
      End If
   Next r_int_Contad
   
   'Confirma
   If MsgBox("¿Está seguro que se abonaron las solicitudes seleccionadas?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   ReDim modatecli_g_arr_TitOpe(0)
    
 
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If (grd_Listad.TextMatrix(r_int_Contad, 14) = "X") Then
         ReDim Preserve modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe) + 1)
         modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_NumOpe = Replace(Trim(grd_Listad.TextMatrix(r_int_Contad, 0)), "-", "")
         modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_FecAct = Right(grd_Listad.TextMatrix(r_int_Contad, 4), 4) & Mid(grd_Listad.TextMatrix(r_int_Contad, 4), 4, 2) & Mid(grd_Listad.TextMatrix(r_int_Contad, 4), 1, 2)
      End If
   Next r_int_Contad
   
   'Actualiza el estado del prepago
   For r_int_Contad = 1 To UBound(modatecli_g_arr_TitOpe)

      g_str_Parame = "USP_ACTUALIZA_CRE_PPGCAB ("
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_TitOpe(r_int_Contad).CreHip_NumOpe & "', "
      g_str_Parame = g_str_Parame & "" & modatecli_g_arr_TitOpe(r_int_Contad).CreHip_FecAct & " , 3, 0, 0, "
      g_str_Parame = g_str_Parame & Format(CDate(Now), "yyyymmdd") & ", 0 ) "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "No se pudo completar la actualización del estado de los datos.", vbInformation, modgen_g_con_PltPar
         Exit Sub
      End If
   Next r_int_Contad
   
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(False)
   Call fs_Buscar
End Sub

Private Sub cmd_Imprim_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer

   'valida selección
   r_int_ConSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 14) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionado Solicitudes para imprimir.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_int_ConSel > 6 Then
      MsgBox "Soló se permite imprimir hasta 6 Solicitudes por Carta.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
    'Valida que las solicitudes seleccionadas solo sean --> Prepago Generado (Est = 1)
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If (grd_Listad.TextMatrix(r_int_Contad, 14) = "X") Then
         If Not fs_Valida_EstPrepago(grd_Listad.TextMatrix(r_int_Contad, 13), 1) Then
            If Trim(grd_Listad.TextMatrix(r_int_Contad, 13)) <> "ENVIADO A COFIDE" Then
               MsgBox "La solicitud " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede imprimir, porque no se encuentra en instancia PREPAGO GENERADO.", vbInformation, modgen_g_str_NomPlt
            Else
               MsgBox "La solicitud " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede imprimir, porque ya se encuentra en instancia ENVIO A COFIDE.", vbInformation, modgen_g_str_NomPlt
            End If
            Screen.MousePointer = 0
            Exit Sub
         End If
      End If
   Next r_int_Contad
   
   'Confirma
   If MsgBox("¿Está seguro de Imprimir las solicitudes seleccionadas?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   ReDim modatecli_g_arr_TitOpe(0)
 
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If (grd_Listad.TextMatrix(r_int_Contad, 14) = "X") Then
         ReDim Preserve modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe) + 1)
         modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_NumOpe = Replace(Trim(grd_Listad.TextMatrix(r_int_Contad, 0)), "-", "")
         modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_FecAct = Right(grd_Listad.TextMatrix(r_int_Contad, 4), 4) & Mid(grd_Listad.TextMatrix(r_int_Contad, 4), 4, 2) & Mid(grd_Listad.TextMatrix(r_int_Contad, 4), 1, 2)
      End If
   Next r_int_Contad
   
   moddat_g_int_FlgAct = 1
   frm_Con_PreSeg_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      'IMPRIMIR
      crp_Imprim.SelectionFormula = ""
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".RPT_TABLA_TEMP"
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_USUCRE} = '" & modgen_g_str_CodUsu & "' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_TERCRE} = '" & modgen_g_str_NombPC & "' AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_NOMBRE} = 'REPORTE PREPAGOS A COFIDE' "
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ope_prepag_01.rpt"
      crp_Imprim.WindowShowPrintSetupBtn = True
      crp_Imprim.Destination = crptToWindow
      crp_Imprim.Action = 1
   End If
   
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_RepCro_Click()
Dim r_int_Contad   As Integer

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
      
   'Valida que las solicitudes seleccionadas solo sean --> Debe tener estado de Enviado a COFIDE (Est = 2) para que pase a Abono a COFIDE
   r_int_Contad = grd_Listad.Row
   If Not fs_Valida_EstPrepago(grd_Listad.TextMatrix(r_int_Contad, 13), 3) Then
      If Trim(grd_Listad.TextMatrix(r_int_Contad, 13)) <> "RECEPCION DE CALENDARIO" Then
         MsgBox "La solicitud " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede Recepcionar, porque no se encuentra en instancia ABONO A COFIDE.", vbInformation, modgen_g_str_NomPlt
      Else
         MsgBox "La solicitud " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede Recepcionar, porque ya se encuentra en instancia RECEPCION DE CALENDARIO.", vbInformation, modgen_g_str_NomPlt
      End If
      
      If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16)) = "" Then
         Exit Sub
      End If
   End If
   
   moddat_g_str_Codigo = ""
   moddat_g_int_FlgGrb = 0
   If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16)) = "" Then
      moddat_g_int_FlgGrb = 1
   End If
   
   frm_Con_PreSeg_03.Show 1
End Sub
                         
Private Sub cmd_RepCro_Click_old()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer
     'valida selección
   r_int_ConSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 14) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad

   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionado Solicitudes para Recepcionar.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Valida que las solicitudes seleccionadas solo sean --> Debe tener estado de Enviado a COFIDE (Est = 2) para que pase a Abono a COFIDE
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If (grd_Listad.TextMatrix(r_int_Contad, 14) = "X") Then
         If Not fs_Valida_EstPrepago(grd_Listad.TextMatrix(r_int_Contad, 13), 3) Then
            If Trim(grd_Listad.TextMatrix(r_int_Contad, 13)) <> "RECEPCION DE CALENDARIO" Then
               MsgBox "La solicitud " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede Recepcionar, porque no se encuentra en instancia ABONO A COFIDE.", vbInformation, modgen_g_str_NomPlt
            Else
               MsgBox "La solicitud " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede Recepcionar, porque ya se encuentra en instancia RECEPCION DE CALENDARIO.", vbInformation, modgen_g_str_NomPlt
            End If
            Screen.MousePointer = 0
            Exit Sub
         End If
      End If
   Next r_int_Contad
   
   'Confirma
   If MsgBox("¿Está seguro que se recepcionaron los calendarios de las solicitudes seleccionadas?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   ReDim modatecli_g_arr_TitOpe(0)
      
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If (grd_Listad.TextMatrix(r_int_Contad, 14) = "X") Then
         ReDim Preserve modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe) + 1)
         modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_NumOpe = Replace(Trim(grd_Listad.TextMatrix(r_int_Contad, 0)), "-", "")
         modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_FecAct = Right(grd_Listad.TextMatrix(r_int_Contad, 4), 4) & Mid(grd_Listad.TextMatrix(r_int_Contad, 4), 4, 2) & Mid(grd_Listad.TextMatrix(r_int_Contad, 4), 1, 2)
         modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_TipCli = Trim(grd_Listad.TextMatrix(r_int_Contad, 3))
         
      End If
   Next r_int_Contad
   
   'Actualiza el estado del prepago
   For r_int_Contad = 1 To UBound(modatecli_g_arr_TitOpe)

      g_str_Parame = "USP_ACTUALIZA_CRE_PPGCAB ("
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_TitOpe(r_int_Contad).CreHip_NumOpe & "', "
      
      If modatecli_g_arr_TitOpe(r_int_Contad).CreHip_TipCli = "TOTAL" Then
         g_str_Parame = g_str_Parame & "" & modatecli_g_arr_TitOpe(r_int_Contad).CreHip_FecAct & " , 5, 0, 0, 0, "
      Else
         g_str_Parame = g_str_Parame & "" & modatecli_g_arr_TitOpe(r_int_Contad).CreHip_FecAct & " , 4, 0, 0, 0, "
      End If
      g_str_Parame = g_str_Parame & Format(CDate(Now), "yyyymmdd") & ") "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "No se pudo completar la actualización del estado de los datos.", vbInformation, modgen_g_con_PltPar
         Exit Sub
      End If
   Next r_int_Contad
   
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
   Call fs_Limpia
   Call fs_Activa(True)
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1225
   grd_Listad.ColWidth(1) = 1210
   grd_Listad.ColWidth(2) = 3970
   grd_Listad.ColWidth(3) = 2230 '2430
   grd_Listad.ColWidth(4) = 1100
   grd_Listad.ColWidth(5) = 1170
   grd_Listad.ColWidth(6) = 1580 '1720
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 2400 '1855 'Estado
   grd_Listad.ColWidth(14) = 1290 '1450 'Seleccionar
   grd_Listad.ColWidth(15) = 0 '1450    'HIPMAE_CODMON
   grd_Listad.ColWidth(16) = 0 'PPGPAS_CODREG
   grd_Listad.ColWidth(17) = 0 'importe
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(14) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmd_Imprim.Enabled = p_Activa
End Sub

Public Sub fs_Buscar()
Dim r_int_FlgEst1 As Integer
Dim r_int_FlgEst2 As Integer
Dim r_int_FlgEst3 As Integer
Dim r_int_FlgEst4 As Integer
Dim r_dbl_ImpAux  As Double

   Call gs_LimpiaGrid(grd_Listad)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PP.PPGCAB_NUMOPE, CH.HIPMAE_TDOCLI, CH.HIPMAE_NDOCLI, CH.HIPMAE_MONEDA, PP.PPGCAB_TIPPPG, "
   g_str_Parame = g_str_Parame & "       PP.PPGCAB_FECPRO, PP.PPGCAB_FECPPG, PP.PPGCAB_MTODEP, PP.PPGCAB_MTOTOT, PP.PPGCAB_TIPPPGPAR, "
   g_str_Parame = g_str_Parame & "       TRIM(CL.DATGEN_APEPAT)||' '||TRIM(CL.DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS CLIENTE, PP.PPGCAB_FLGEST, "
   g_str_Parame = g_str_Parame & "       (SELECT X.PPGPAS_CODREG FROM CRE_PPGPASCTB X  "
   g_str_Parame = g_str_Parame & "                              WHERE x.PPGPAS_NUMOPE = PP.PPGCAB_NUMOPE And x.PPGPAS_FECPPG = PP.PPGCAB_FECPPG  "
   g_str_Parame = g_str_Parame & "                                AND X.PPGPAS_SITUAC = 1) CODREG  "
   g_str_Parame = g_str_Parame & "  FROM CRE_PPGCAB PP "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_HIPMAE CH ON CH.HIPMAE_NUMOPE = PP.PPGCAB_NUMOPE "
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN CL ON CL.DATGEN_TIPDOC = CH.HIPMAE_TDOCLI AND CL.DATGEN_NUMDOC = CH.HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & " WHERE (PP.PPGCAB_FLGEST = 1 OR PP.PPGCAB_FLGEST = 2 OR PP.PPGCAB_FLGEST = 3 OR PP.PPGCAB_FLGEST = 4) "
   g_str_Parame = g_str_Parame & "   AND SUBSTR(PP.PPGCAB_NUMOPE,1,3) NOT IN ('001','002','006','011')"
   g_str_Parame = g_str_Parame & " ORDER BY PP.PPGCAB_NUMOPE ASC, PP.PPGCAB_FECPPG ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, l_rst_Prepagos, 3) Then
      Exit Sub
   End If

   If l_rst_Prepagos.BOF And l_rst_Prepagos.EOF Then
      l_rst_Prepagos.Close
      Set l_rst_Prepagos = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   l_rst_Prepagos.MoveFirst
   Do While Not l_rst_Prepagos.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      r_dbl_ImpAux = 0
      
      'numero operacion (formateado)
      grd_Listad.Col = 0
      grd_Listad.Text = gf_Formato_NumOpe(Trim(l_rst_Prepagos!PPGCAB_NUMOPE & ""))
      
      'tipo de documento
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(l_rst_Prepagos!HIPMAE_TDOCLI) & "-" & Trim(l_rst_Prepagos!HIPMAE_NDOCLI & "")
      
      'nombre del cliente
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(l_rst_Prepagos!CLIENTE & "")
      
      'tipo de prepago
      grd_Listad.Col = 3
      If l_rst_Prepagos!PPGCAB_TIPPPG = 1 Then
         If l_rst_Prepagos!PPGCAB_TIPPPGPAR = 1 Then
            grd_Listad.Text = "PARCIAL - RED MONTO"
         Else
            grd_Listad.Text = "PARCIAL - RED PLAZO"
         End If
      Else
        grd_Listad.Text = "TOTAL"
      End If
      
      'fecha del prepago (formateado)
      grd_Listad.Col = 4
      grd_Listad.Text = gf_FormatoFecha(CStr(l_rst_Prepagos!PPGCAB_FECPPG))
      
      'fecha de proceso (formateado)
      grd_Listad.Col = 5
      grd_Listad.Text = gf_FormatoFecha(CStr(l_rst_Prepagos!PPGCAB_FECPRO))
      
      'importe del prepago (formateado)
      grd_Listad.Col = 6
      If l_rst_Prepagos!PPGCAB_TIPPPG = 1 Then
         If l_rst_Prepagos!HIPMAE_MONEDA = 1 Then
            grd_Listad.Text = "S/.   " & Format(l_rst_Prepagos!PPGCAB_MTODEP, "###,###,###,##0.00")
         Else
            grd_Listad.Text = "US$   " & Format(l_rst_Prepagos!PPGCAB_MTODEP, "###,###,###,##0.00")
         End If
         r_dbl_ImpAux = l_rst_Prepagos!PPGCAB_MTODEP
      Else
         If l_rst_Prepagos!HIPMAE_MONEDA = 1 Then
            grd_Listad.Text = "S/.   " & Format(l_rst_Prepagos!PPGCAB_MTOTOT, "###,###,###,##0.00")
         Else
            grd_Listad.Text = "US$   " & Format(l_rst_Prepagos!PPGCAB_MTOTOT, "###,###,###,##0.00")
         End If
         r_dbl_ImpAux = l_rst_Prepagos!PPGCAB_MTOTOT
      End If
      
      ' Tipo de prepago (parcial o total)
      grd_Listad.Col = 7
      grd_Listad.Text = Trim(l_rst_Prepagos!PPGCAB_TIPPPG)
      
      ' Tipo de prepago Parcial (monto o tiempo)
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(l_rst_Prepagos!PPGCAB_TIPPPGPAR & "")
      
      'numero de operacion
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(l_rst_Prepagos!PPGCAB_NUMOPE & "")
      
      'fecha de prepago
      grd_Listad.Col = 10
      grd_Listad.Text = CStr(l_rst_Prepagos!PPGCAB_FECPPG)
      
      'fecha de proceso
      grd_Listad.Col = 11
      grd_Listad.Text = CStr(l_rst_Prepagos!PPGCAB_FECPRO)
      
      'importe del prepago
      grd_Listad.Col = 12
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_MTOTOT
      
      'Estado del prepago
      grd_Listad.Col = 13
      grd_Listad.Text = moddat_gf_Consulta_ParDes("523", l_rst_Prepagos!PPGCAB_FLGEST)
      If l_rst_Prepagos!PPGCAB_FLGEST = 1 Then
         grd_Listad.CellForeColor = modgen_g_con_ColRoj
         r_int_FlgEst1 = r_int_FlgEst1 + 1
      ElseIf l_rst_Prepagos!PPGCAB_FLGEST = 2 Then
         r_int_FlgEst2 = r_int_FlgEst2 + 1
         grd_Listad.CellForeColor = modgen_g_con_ColAzu
      ElseIf l_rst_Prepagos!PPGCAB_FLGEST = 3 Then
         r_int_FlgEst3 = r_int_FlgEst3 + 1
         grd_Listad.CellForeColor = modgen_g_con_ColVer
      ElseIf l_rst_Prepagos!PPGCAB_FLGEST = 4 Then
         r_int_FlgEst4 = r_int_FlgEst4 + 1
         grd_Listad.CellForeColor = modgen_g_con_ColNar
      End If
      
      'CODIGO DE MONEDA
      grd_Listad.Col = 15
      grd_Listad.Text = l_rst_Prepagos!HIPMAE_MONEDA
      
      grd_Listad.Col = 16
      grd_Listad.Text = Trim(l_rst_Prepagos!CODREG & "")
      
      grd_Listad.Col = 17
      grd_Listad.Text = r_dbl_ImpAux

      l_rst_Prepagos.MoveNext
   Loop
   grd_Listad.Redraw = True
   Call fs_Activa(True)
   
   If grd_Listad.Rows > 0 Then
      grd_Listad.Enabled = True
   End If
   
   pnl_PPGRea.Caption = r_int_FlgEst1 & " "
   pnl_CarCOF.Caption = r_int_FlgEst2 & " "
   pnl_AboCOF.Caption = r_int_FlgEst3 & " "
   pnl_RecCro.Caption = r_int_FlgEst4 & " "
      
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 14) = ""
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 14) = "X"
         Next r_Fila
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Public Function fs_Valida_EstPrepago(ByVal p_EstPrePag As String, ByVal p_CodEst As Integer) As Boolean
   fs_Valida_EstPrepago = False
   If p_CodEst = 1 Then
      If InStr(p_EstPrePag, "ENVIADO") = 0 And InStr(p_EstPrePag, "ABONO") = 0 And InStr(p_EstPrePag, "RECEPCION") = 0 Then
         fs_Valida_EstPrepago = True
      End If
   ElseIf p_CodEst = 2 Then
      If InStr(p_EstPrePag, "ENVIADO") > 0 Then
         fs_Valida_EstPrepago = True
      End If
   ElseIf p_CodEst = 3 Then
      If InStr(p_EstPrePag, "ABONO") > 0 Then
         fs_Valida_EstPrepago = True
      End If
   End If
End Function

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      If grd_Listad.TextMatrix(grd_Listad.Row, 14) = "X" Then
         grd_Listad.TextMatrix(grd_Listad.Row, 14) = ""
      Else
         grd_Listad.TextMatrix(grd_Listad.Row, 14) = "X"
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub cmd_Export_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "SEGUIMIENTO DE PREPAGOS" & "  (" & Format(date, "DD/MM/YYYY") & ")"
      .Range(.Cells(2, 2), .Cells(2, 10)).Merge
      .Range(.Cells(2, 2), .Cells(2, 10)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 10)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(4, 2) = "NRO OPERACION"
      .Cells(4, 3) = "ID CLIENTE"
      .Cells(4, 4) = "NOMBRE CLIENTE"
      .Cells(4, 5) = "TIPO DE PREPAGO"
      .Cells(4, 6) = "FECHA PREPAGO"
      .Cells(4, 7) = "FECHA PROCESO"
      .Cells(4, 8) = "MONEDA"
      .Cells(4, 9) = "IMPORTE"
      .Cells(4, 10) = "ESTADO"
      
      .Range(.Cells(4, 2), .Cells(4, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 10)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 17
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 13
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 43
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 25
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 17
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 16
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 14
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 16
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 30
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Size = 11
      
      r_int_NumFil = 2
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil + 3, 2) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 2, 0) 'NRO OPERACION
         .Cells(r_int_NumFil + 3, 3) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 2, 1) 'ID-CLIENTE
         .Cells(r_int_NumFil + 3, 4) = grd_Listad.TextMatrix(r_int_NumFil - 2, 2) 'NOMBRE CLIENTE
         .Cells(r_int_NumFil + 3, 5) = grd_Listad.TextMatrix(r_int_NumFil - 2, 3) 'TIPO DE PREPAGO
         .Cells(r_int_NumFil + 3, 6) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 2, 4) 'FECHA PREPAGO
         .Cells(r_int_NumFil + 3, 7) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 2, 5) 'FECHA PROCESO
         .Cells(r_int_NumFil + 3, 8) = moddat_gf_Consulta_ParDes("204", CStr(grd_Listad.TextMatrix(r_int_NumFil - 2, 15))) 'MONEDA
         .Cells(r_int_NumFil + 3, 9) = grd_Listad.TextMatrix(r_int_NumFil - 2, 17) 'IMPORTE
         .Cells(r_int_NumFil + 3, 10) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 2, 13) 'ESTADO
         
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
      
      .Range(.Cells(4, 4), .Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

'Reordenar
Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_DoiCli_Click()
   If Len(Trim(pnl_Tit_DoiCli.Tag)) = 0 Or pnl_Tit_DoiCli.Tag = "D" Then
      pnl_Tit_DoiCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_DoiCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
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

Private Sub pnl_Tit_TipPpg_Click()
   If Len(Trim(pnl_Tit_TipPpg.Tag)) = 0 Or pnl_Tit_TipPpg.Tag = "D" Then
      pnl_Tit_TipPpg.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_TipPpg.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecPpg_Click()
   If Len(Trim(pnl_Tit_FecPpg.Tag)) = 0 Or pnl_Tit_FecPpg.Tag = "D" Then
      pnl_Tit_FecPpg.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 10, "C")
   Else
      pnl_Tit_FecPpg.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 10, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecPro_Click()
   If Len(Trim(pnl_Tit_FecPro.Tag)) = 0 Or pnl_Tit_FecPro.Tag = "D" Then
      pnl_Tit_FecPro.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 11, "C")
   Else
      pnl_Tit_FecPro.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 11, "C-")
   End If
End Sub

Private Sub pnl_Tit_Import_Click()
   If Len(Trim(pnl_Tit_Import.Tag)) = 0 Or pnl_Tit_Import.Tag = "D" Then
      pnl_Tit_Import.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 12, "N")
   Else
      pnl_Tit_Import.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 12, "N-")
   End If
End Sub

Private Sub pnl_Tit_Estado_Click()
   If Len(Trim(pnl_Tit_Estado.Tag)) = 0 Or pnl_Tit_Estado.Tag = "D" Then
      pnl_Tit_Estado.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 13, "C")
   Else
      pnl_Tit_Estado.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 13, "C-")
   End If
End Sub
