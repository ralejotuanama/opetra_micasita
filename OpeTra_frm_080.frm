VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_ConCre_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10200
   ClientLeft      =   -510
   ClientTop       =   780
   ClientWidth     =   15420
   Icon            =   "OpeTra_frm_080.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16965
      _Version        =   65536
      _ExtentX        =   29924
      _ExtentY        =   17965
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   44
         Top             =   9690
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   315
            Left            =   60
            TabIndex        =   45
            Top             =   60
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   15
            BackColor       =   16711680
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
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   315
            Left            =   4470
            TabIndex        =   47
            Top             =   60
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   15
            BackColor       =   8421376
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
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   315
            Left            =   9360
            TabIndex        =   49
            Top             =   60
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   15
            BackColor       =   192
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
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   315
            Left            =   13200
            TabIndex        =   51
            Top             =   60
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   15
            BackColor       =   0
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
            Outline         =   -1  'True
         End
         Begin VB.Label Label3 
            Caption         =   "Cuotas Pendientes de Pago"
            Height          =   315
            Left            =   14100
            TabIndex        =   52
            Top             =   90
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "Cuotas Atrasadas"
            Height          =   315
            Left            =   10260
            TabIndex        =   50
            Top             =   90
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "Cuotas Pagadas después del Vencimiento"
            Height          =   315
            Left            =   5370
            TabIndex        =   48
            Top             =   90
            Width           =   3075
         End
         Begin VB.Label lbl_Descri 
            Caption         =   "Cuotas Pagadas antes del Vencimiento"
            Height          =   315
            Left            =   960
            TabIndex        =   46
            Top             =   90
            Width           =   3075
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   750
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
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
            Left            =   16260
            Picture         =   "OpeTra_frm_080.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
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
            Left            =   630
            TabIndex        =   4
            Top             =   60
            Width           =   10095
            _Version        =   65536
            _ExtentX        =   17806
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Consulta de Crédito Hipotecario - Cronograma de Pagos"
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
            Picture         =   "OpeTra_frm_080.frx":044E
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   435
         Left            =   30
         TabIndex        =   5
         Top             =   1440
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1560
            TabIndex        =   6
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   6900
            TabIndex        =   7
            Top             =   60
            Width           =   9915
            _Version        =   65536
            _ExtentX        =   17489
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   6090
            TabIndex        =   8
            Top             =   60
            Width           =   735
         End
      End
      Begin Threed.SSPanel SSPanel78 
         Height          =   7275
         Left            =   30
         TabIndex        =   10
         Top             =   1920
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
         _ExtentY        =   12832
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
            Height          =   285
            Left            =   7050
            TabIndex        =   11
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cap. PBP"
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
         Begin MSFlexGridLib.MSFlexGrid grd_CliNCo_Listad 
            Height          =   6855
            Left            =   30
            TabIndex        =   12
            Top             =   360
            Width           =   16815
            _ExtentX        =   29660
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   30
            Cols            =   17
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   2550
            TabIndex        =   13
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Interés"
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
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1032
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuota"
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
         Begin Threed.SSPanel SSPanel38 
            Height          =   285
            Left            =   630
            TabIndex        =   15
            Top             =   60
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vcto"
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
         Begin Threed.SSPanel SSPanel41 
            Height          =   285
            Left            =   1650
            TabIndex        =   16
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Amortiz."
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
         Begin Threed.SSPanel SSPanel42 
            Height          =   285
            Left            =   12420
            TabIndex        =   17
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Cuota"
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
         Begin Threed.SSPanel SSPanel43 
            Height          =   285
            Left            =   13320
            TabIndex        =   18
            Top             =   60
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "S. Capital"
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
         Begin Threed.SSPanel SSPanel58 
            Height          =   285
            Left            =   3450
            TabIndex        =   19
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Sg. Prest."
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
         Begin Threed.SSPanel SSPanel60 
            Height          =   285
            Left            =   4350
            TabIndex        =   20
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Sg. Inm."
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
         Begin Threed.SSPanel SSPanel63 
            Height          =   285
            Left            =   5250
            TabIndex        =   21
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Portes"
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
         Begin Threed.SSPanel SSPanel56 
            Height          =   285
            Left            =   7920
            TabIndex        =   22
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Int. PBP"
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
         Begin Threed.SSPanel SSPanel65 
            Height          =   285
            Left            =   14400
            TabIndex        =   23
            Top             =   60
            Width           =   2100
            _Version        =   65536
            _ExtentX        =   3704
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
         Begin Threed.SSPanel SSPanel77 
            Height          =   285
            Left            =   8820
            TabIndex        =   24
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Int Morat."
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
         Begin Threed.SSPanel SSPanel79 
            Height          =   285
            Left            =   9720
            TabIndex        =   25
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Int Comp."
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
         Begin Threed.SSPanel SSPanel80 
            Height          =   285
            Left            =   10620
            TabIndex        =   26
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "G. Cob."
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
            Left            =   11520
            TabIndex        =   27
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Otr. Gastos"
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
            Left            =   6150
            TabIndex        =   38
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Cuota"
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
      Begin Threed.SSPanel SSPanel64 
         Height          =   405
         Left            =   30
         TabIndex        =   28
         Top             =   9240
         Width           =   16875
         _Version        =   65536
         _ExtentX        =   29766
         _ExtentY        =   714
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
         Begin Threed.SSPanel pnl_Capita 
            Height          =   285
            Left            =   1650
            TabIndex        =   29
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_SegViv 
            Height          =   285
            Left            =   4350
            TabIndex        =   30
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_SegPre 
            Height          =   285
            Left            =   3450
            TabIndex        =   31
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Intere 
            Height          =   285
            Left            =   2550
            TabIndex        =   32
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Portes 
            Height          =   285
            Left            =   5250
            TabIndex        =   33
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_SubTot 
            Height          =   285
            Left            =   6150
            TabIndex        =   34
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CapBBP 
            Height          =   285
            Left            =   7050
            TabIndex        =   35
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_IntBBP 
            Height          =   285
            Left            =   7950
            TabIndex        =   36
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_IntMor 
            Height          =   285
            Left            =   8850
            TabIndex        =   39
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_IntCom 
            Height          =   285
            Left            =   9750
            TabIndex        =   40
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_GasCob 
            Height          =   285
            Left            =   10650
            TabIndex        =   41
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_OtrGas 
            Height          =   285
            Left            =   11550
            TabIndex        =   42
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotCuo 
            Height          =   285
            Left            =   12450
            TabIndex        =   43
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin VB.Label lbl_Totale 
            Alignment       =   1  'Right Justify
            Caption         =   "Totales ===> US$ "
            Height          =   315
            Index           =   0
            Left            =   30
            TabIndex        =   37
            Top             =   30
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_ConCre_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   
   Call fs_Carga_Cro_CliNCo
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Cliente No Concesional
   grd_CliNCo_Listad.ColWidth(0) = 575
   grd_CliNCo_Listad.ColWidth(1) = 1025
   grd_CliNCo_Listad.ColWidth(2) = 900
   grd_CliNCo_Listad.ColWidth(3) = 900
   grd_CliNCo_Listad.ColWidth(4) = 900
   grd_CliNCo_Listad.ColWidth(5) = 900
   grd_CliNCo_Listad.ColWidth(6) = 900
   grd_CliNCo_Listad.ColWidth(7) = 900
   grd_CliNCo_Listad.ColWidth(8) = 900
   grd_CliNCo_Listad.ColWidth(9) = 900
   grd_CliNCo_Listad.ColWidth(10) = 890
   grd_CliNCo_Listad.ColWidth(11) = 890
   grd_CliNCo_Listad.ColWidth(12) = 890
   grd_CliNCo_Listad.ColWidth(13) = 900
   grd_CliNCo_Listad.ColWidth(14) = 900
   grd_CliNCo_Listad.ColWidth(15) = 1080
   grd_CliNCo_Listad.ColWidth(16) = 2100
   
   grd_CliNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(10) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(11) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(12) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(13) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(14) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(15) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(16) = flexAlignCenterCenter
End Sub

Private Sub fs_Carga_Cro_CliNCo()
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_SegDes     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_OtrCar     As Double
   Dim r_dbl_SubTot     As Double
   Dim r_dbl_CapBBP     As Double
   Dim r_dbl_IntBBP     As Double
   Dim r_dbl_IntMor     As Double
   Dim r_dbl_IntCom     As Double
   Dim r_dbl_GasCob     As Double
   Dim r_dbl_OtrGas     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   Dim r_var_ColCel
   Dim r_str_Situac     As String
   
   Call gs_LimpiaGrid(grd_CliNCo_Listad)

   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_SegDes = 0
   r_dbl_SegViv = 0
   r_dbl_OtrCar = 0
   r_dbl_SubTot = 0
   r_dbl_CapBBP = 0
   r_dbl_IntBBP = 0
   r_dbl_IntMor = 0
   r_dbl_IntCom = 0
   r_dbl_GasCob = 0
   r_dbl_OtrGas = 0
   r_dbl_TotCuo = 0

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_CliNCo_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_CliNCo_Listad.Rows = grd_CliNCo_Listad.Rows + 1
         grd_CliNCo_Listad.Row = grd_CliNCo_Listad.Rows - 1
         
         
         r_str_Situac = ""
         If g_rst_Princi!HIPCUO_SITUAC = 1 Then
            If CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))) <= CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))) Then
               r_var_ColCel = modgen_g_con_ColAzu
            Else
               r_var_ColCel = modgen_g_con_ColCya
            End If
         Else
            If CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))) < Date Then
               r_var_ColCel = modgen_g_con_ColRoj
               r_str_Situac = "ATRASADO"
            Else
               r_var_ColCel = modgen_g_con_ColNeg
            End If
         End If
         
         grd_CliNCo_Listad.Col = 16
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         
         If r_str_Situac <> "ATRASADO" Then
            grd_CliNCo_Listad.Text = moddat_gf_Consulta_ParDes("001", CStr(g_rst_Princi!HIPCUO_SITUAC))
         Else
            grd_CliNCo_Listad.Text = r_str_Situac
         End If
         
         r_dbl_ImpCuo = 0
         
         grd_CliNCo_Listad.Col = 0
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_CliNCo_Listad.Col = 1
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_CliNCo_Listad.Col = 2
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 3
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 4
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 5
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 6
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 7
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         r_dbl_SubTot = r_dbl_SubTot + r_dbl_ImpCuo
         
         grd_CliNCo_Listad.Col = 8
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPBBP, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 9
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_INTBBP, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 10
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_INTMOR, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 11
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_INTCOM, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 12
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_GASCOB, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 13
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_OTRGAS, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 14
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_CliNCo_Listad.Col = 15
         grd_CliNCo_Listad.CellForeColor = r_var_ColCel
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")


         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_SegDes = r_dbl_SegDes + CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_SegViv = r_dbl_SegViv + CDbl(Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00"))
         r_dbl_OtrCar = r_dbl_OtrCar + CDbl(Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00"))
         r_dbl_CapBBP = r_dbl_CapBBP + CDbl(Format(g_rst_Princi!HIPCUO_CAPBBP, "###,###,##0.00"))
         r_dbl_IntBBP = r_dbl_IntBBP + CDbl(Format(g_rst_Princi!HIPCUO_INTBBP, "###,###,##0.00"))
         r_dbl_IntMor = r_dbl_IntMor + CDbl(Format(g_rst_Princi!HIPCUO_INTMOR, "###,###,##0.00"))
         r_dbl_IntCom = r_dbl_IntCom + CDbl(Format(g_rst_Princi!HIPCUO_INTCOM, "###,###,##0.00"))
         r_dbl_GasCob = r_dbl_GasCob + CDbl(Format(g_rst_Princi!HIPCUO_GASCOB, "###,###,##0.00"))
         r_dbl_OtrGas = r_dbl_OtrGas + CDbl(Format(g_rst_Princi!HIPCUO_OTRGAS, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
            
         g_rst_Princi.MoveNext
      Loop
      
      grd_CliNCo_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_CliNCo_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_SegPre.Caption = Format(r_dbl_SegDes, "###,###,##0.00") & " "
   pnl_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
   pnl_Portes.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
   pnl_SubTot.Caption = Format(r_dbl_SubTot, "###,###,##0.00") & " "
   pnl_CapBBP.Caption = Format(r_dbl_CapBBP, "###,###,##0.00") & " "
   pnl_IntBBP.Caption = Format(r_dbl_IntBBP, "###,###,##0.00") & " "
   pnl_IntMor.Caption = Format(r_dbl_IntMor, "###,###,##0.00") & " "
   pnl_IntCom.Caption = Format(r_dbl_IntCom, "###,###,##0.00") & " "
   pnl_GasCob.Caption = Format(r_dbl_GasCob, "###,###,##0.00") & " "
   pnl_OtrGas.Caption = Format(r_dbl_OtrGas, "###,###,##0.00") & " "
   pnl_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub

Private Sub grd_CliNCo_Listad_DblClick()
   If grd_CliNCo_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_CliNCo_Listad.Col = 0
   moddat_g_int_NumCuo = CInt(grd_CliNCo_Listad)
   
   Call gs_RefrescaGrid(grd_CliNCo_Listad)
   
   frm_ConCre_04.Show 1
End Sub

Private Sub grd_CliNCo_Listad_SelChange()
   If grd_CliNCo_Listad.Rows > 2 Then
      grd_CliNCo_Listad.RowSel = grd_CliNCo_Listad.Row
   End If
End Sub

