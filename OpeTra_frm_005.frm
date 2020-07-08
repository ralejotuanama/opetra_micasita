VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Cob_MovDia_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   2865
   ClientTop       =   1950
   ClientWidth     =   7905
   Icon            =   "OpeTra_frm_005.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      _Version        =   65536
      _ExtentX        =   13944
      _ExtentY        =   13309
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
         Height          =   2775
         Left            =   30
         TabIndex        =   1
         Top             =   3900
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   4895
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
            Left            =   1860
            TabIndex        =   2
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_CodIte 
            Height          =   315
            Left            =   1860
            TabIndex        =   3
            Top             =   390
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_TipDoc 
            Height          =   315
            Left            =   1860
            TabIndex        =   4
            Top             =   720
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1860
            TabIndex        =   5
            Top             =   1050
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   1860
            TabIndex        =   6
            Top             =   1380
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_Import 
            Height          =   315
            Left            =   1860
            TabIndex        =   7
            Top             =   1710
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ImpITF 
            Height          =   315
            Left            =   1860
            TabIndex        =   8
            Top             =   2040
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ImpTot 
            Height          =   315
            Left            =   1860
            TabIndex        =   9
            Top             =   2370
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin VB.Label Label5 
            Caption         =   "Nro. Operac./Refer.:"
            Height          =   315
            Left            =   90
            TabIndex        =   17
            Top             =   60
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Código Item:"
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Doc. Ide. Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Nombre Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   14
            Top             =   1050
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   90
            TabIndex        =   13
            Top             =   1380
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Importe:"
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   1710
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "ITF:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   2040
            Width           =   1155
         End
         Begin VB.Label Label12 
            Caption         =   "Importe Total:"
            Height          =   315
            Left            =   90
            TabIndex        =   10
            Top             =   2370
            Width           =   1155
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3075
         Left            =   30
         TabIndex        =   18
         Top             =   780
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   5424
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
         Begin Threed.SSPanel pnl_NomUsu 
            Height          =   315
            Left            =   1860
            TabIndex        =   19
            Top             =   1380
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_HorMov 
            Height          =   315
            Left            =   1860
            TabIndex        =   20
            Top             =   390
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_NumMov 
            Height          =   315
            Left            =   1860
            TabIndex        =   21
            Top             =   720
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_TipMov 
            Height          =   315
            Left            =   1860
            TabIndex        =   22
            Top             =   1050
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_FlgRev 
            Height          =   315
            Left            =   5310
            TabIndex        =   23
            Top             =   90
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "OPERACION REVERSADA"
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
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_CodBan 
            Height          =   315
            Left            =   1860
            TabIndex        =   32
            Top             =   1710
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_FecMov 
            Height          =   315
            Left            =   1860
            TabIndex        =   34
            Top             =   60
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_NumCta 
            Height          =   315
            Left            =   1860
            TabIndex        =   36
            Top             =   2040
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_FecPag 
            Height          =   315
            Left            =   1860
            TabIndex        =   38
            Top             =   2370
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_NumCom 
            Height          =   315
            Left            =   1860
            TabIndex        =   40
            Top             =   2700
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin VB.Label Label17 
            Caption         =   "Nro. Comprobante:"
            Height          =   255
            Left            =   90
            TabIndex        =   41
            Top             =   2700
            Width           =   1635
         End
         Begin VB.Label Label16 
            Caption         =   "Fecha de Pago:"
            Height          =   315
            Left            =   90
            TabIndex        =   39
            Top             =   2370
            Width           =   1425
         End
         Begin VB.Label Label15 
            Caption         =   "Nro. Cuenta:"
            Height          =   315
            Left            =   90
            TabIndex        =   37
            Top             =   2040
            Width           =   1785
         End
         Begin VB.Label Label14 
            Caption         =   "Fecha de Movimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   35
            Top             =   60
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   90
            TabIndex        =   33
            Top             =   1710
            Width           =   1785
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario:"
            Height          =   315
            Left            =   90
            TabIndex        =   27
            Top             =   1380
            Width           =   1785
         End
         Begin VB.Label Label2 
            Caption         =   "Hora de Movimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   26
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. Movimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   25
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo Movimiento"
            Height          =   315
            Left            =   90
            TabIndex        =   24
            Top             =   1050
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   30
         TabIndex        =   28
         Top             =   30
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   585
            Left            =   630
            TabIndex        =   29
            Top             =   30
            Width           =   6675
            _Version        =   65536
            _ExtentX        =   11774
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Operaciones por Bancos - Consulta de Movimiento"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            Picture         =   "OpeTra_frm_005.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel15 
         Height          =   765
         Left            =   30
         TabIndex        =   30
         Top             =   6720
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
         Begin VB.CommandButton cmd_Revers 
            Height          =   675
            Left            =   6390
            Picture         =   "OpeTra_frm_005.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Reversar Operación"
            Top             =   30
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CommandButton cmd_ImpCom 
            Height          =   675
            Left            =   30
            Picture         =   "OpeTra_frm_005.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Imprimir Comprobante"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7080
            Picture         =   "OpeTra_frm_005.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   2220
            Top             =   180
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
End
Attribute VB_Name = "frm_Cob_MovDia_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Revers_Click()
   Dim r_lng_NueMov     As Long
   
   If MsgBox("¿Está seguro de reversar la operación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Select Case opecaj_g_arr_OpeCaj(1).MovCaj_TipMov
      Case 1101
         If opecaj_rev_Efe_GasAdm(opecaj_g_arr_OpeCaj, opecaj_g_str_UsuMov, "000000", "", opecaj_g_str_NumMov, r_lng_NueMov) Then
            MsgBox "Operación reversada con éxito", vbInformation, modgen_g_str_NomPlt
            
            opecaj_g_int_FlgAct = 2
            Unload Me
         End If
   End Select
End Sub

Private Sub cmd_ImpCom_Click()
   'On Error GoTo Error_Imp

   'dlg_Guarda.CancelError = True
   'dlg_Guarda.ShowPrinter


   Select Case opecaj_g_arr_OpeCaj(1).MovCaj_TipMov
      Case 1101      'Gastos Administrativos
         Call opecaj_gs_Imp_GasAdm_Ban(Format(CDate(opecaj_g_arr_OpeCaj(1).MovCaj_FecMov), "yyyymmdd"), opecaj_g_str_NumMov)
      
      Case 1102      'Cuotas Credito Hipotecario
            Call opecaj_gs_Imp_CuoHip_Ban(Trim(opecaj_g_arr_OpeCaj(1).MovCaj_NumOpe), Format(CDate(opecaj_g_arr_OpeCaj(1).MovCaj_FecMov), "yyyymmdd"), opecaj_g_str_NumMov)
            
   End Select
   
   Call gs_Imprim_ComPag
   
Error_Imp:
   Exit Sub
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_str_HorMov     As String
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call opecaj_gs_Consulta_CajMov(opecaj_g_arr_OpeCaj, opecaj_g_str_UsuMov, opecaj_g_str_CodBan, opecaj_g_str_FecMov, opecaj_g_str_NumMov)
   
   pnl_CodBan.Caption = moddat_gf_Consulta_ParDes("505", opecaj_g_arr_OpeCaj(1).MovCaj_CodBan)
   pnl_NomUsu.Caption = opecaj_g_arr_OpeCaj(1).MovCaj_UsuMov
   
   r_str_HorMov = Format(opecaj_g_arr_OpeCaj(1).MovCaj_HorMov, "000000")
   r_str_HorMov = Mid(r_str_HorMov, 1, 2) & ":" & Mid(r_str_HorMov, 3, 2) & ":" & Mid(r_str_HorMov, 5, 2)

   pnl_FecMov.Caption = opecaj_g_arr_OpeCaj(1).MovCaj_FecMov
   pnl_HorMov.Caption = r_str_HorMov
   pnl_NumMov.Caption = opecaj_g_str_NumMov
   pnl_TipMov.Caption = CStr(opecaj_g_arr_OpeCaj(1).MovCaj_TipMov) & " - " & moddat_gf_Consulta_ParDes("301", opecaj_g_arr_OpeCaj(1).MovCaj_TipMov)
   
   pnl_FecPag.Caption = opecaj_g_arr_OpeCaj(1).MovCaj_FecDep
   pnl_NumCta.Caption = opecaj_g_arr_OpeCaj(1).MovCaj_NumCta
   pnl_NumCom.Caption = opecaj_g_arr_OpeCaj(1).MovCaj_NumCom
   
   pnl_NumOpe.Caption = opecaj_g_arr_OpeCaj(1).MovCaj_NumOpe
   pnl_CodIte.Caption = opecaj_g_arr_OpeCaj(1).MovCaj_CodIte
   
   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("203", CStr(opecaj_g_arr_OpeCaj(1).MovCaj_TipDoc)) & " - " & opecaj_g_arr_OpeCaj(1).MovCaj_NumDoc
   pnl_NomCli.Caption = moddat_gf_Buscar_NomCli(opecaj_g_arr_OpeCaj(1).MovCaj_TipDoc, opecaj_g_arr_OpeCaj(1).MovCaj_NumDoc)

   pnl_Moneda.Caption = moddat_gf_Consulta_ParDes("204", CStr(opecaj_g_arr_OpeCaj(1).MovCaj_MonPag))
   pnl_Import.Caption = Format(opecaj_g_arr_OpeCaj(1).MovCaj_ImpPag, "###,###,##0.00") & " "
   pnl_ImpITF.Caption = Format(opecaj_g_arr_OpeCaj(1).MovCaj_ITFImp, "###,###,##0.00") & " "
   pnl_ImpTot.Caption = Format(opecaj_g_arr_OpeCaj(1).MovCaj_ImpTot, "###,###,##0.00") & " "
   
   If opecaj_g_arr_OpeCaj(1).MovCaj_FlgRev = 0 Then
      pnl_FlgRev.Visible = False
   Else
      pnl_FlgRev.Visible = True
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub


