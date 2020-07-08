VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Ges_TecPro_07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12795
   Icon            =   "OpeTra_frm_829.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8025
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   12795
      _Version        =   65536
      _ExtentX        =   22569
      _ExtentY        =   14155
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
      Begin Threed.SSPanel SSPanel18 
         Height          =   5280
         Left            =   60
         TabIndex        =   1
         Top             =   2670
         Width           =   12675
         _Version        =   65536
         _ExtentX        =   22357
         _ExtentY        =   9313
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
         Begin Threed.SSPanel pnl_Total7 
            Height          =   285
            Left            =   10020
            TabIndex        =   2
            Top             =   2115
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin Threed.SSPanel pnl_Total6 
            Height          =   285
            Left            =   8625
            TabIndex        =   3
            Top             =   2115
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin Threed.SSPanel pnl_Total1 
            Height          =   285
            Left            =   1650
            TabIndex        =   4
            Top             =   2115
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin Threed.SSPanel pnl_Total2 
            Height          =   285
            Left            =   3045
            TabIndex        =   5
            Top             =   2115
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin Threed.SSPanel pnl_Total3 
            Height          =   285
            Left            =   4440
            TabIndex        =   6
            Top             =   2115
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin Threed.SSPanel pnl_Total4 
            Height          =   285
            Left            =   5835
            TabIndex        =   7
            Top             =   2115
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin Threed.SSPanel pnl_Total5 
            Height          =   285
            Left            =   7230
            TabIndex        =   8
            Top             =   2115
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_rcc2 
            Height          =   2280
            Left            =   45
            TabIndex        =   9
            Top             =   2940
            Width           =   12465
            _ExtentX        =   21987
            _ExtentY        =   4022
            _Version        =   393216
            Rows            =   12
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_MtoSol 
            Height          =   285
            Left            =   9120
            TabIndex        =   10
            Top             =   2685
            Width           =   1020
            _Version        =   65536
            _ExtentX        =   1799
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Monto (S/.)"
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
         Begin Threed.SSPanel pnl_TipCla 
            Height          =   285
            Left            =   4170
            TabIndex        =   11
            Top             =   2685
            Width           =   1620
            _Version        =   65536
            _ExtentX        =   2857
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Clasificacion"
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
         Begin Threed.SSPanel pnl_MtoDol 
            Height          =   285
            Left            =   10140
            TabIndex        =   12
            Top             =   2685
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1817
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Monto (US$)"
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
         Begin Threed.SSPanel pnl_TipDeu 
            Height          =   285
            Left            =   5790
            TabIndex        =   13
            Top             =   2685
            Width           =   3330
            _Version        =   65536
            _ExtentX        =   5874
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Deuda"
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_rcc1 
            Height          =   1545
            Left            =   45
            TabIndex        =   14
            Top             =   570
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   2725
            _Version        =   393216
            Rows            =   1
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
         End
         Begin Threed.SSPanel pnl_Periodo1 
            Height          =   285
            Left            =   1650
            TabIndex        =   15
            Top             =   315
            Width           =   1405
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Periodo 1"
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
         Begin Threed.SSPanel pnl_Periodo5 
            Height          =   285
            Left            =   7230
            TabIndex        =   16
            Top             =   315
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Periodo 5"
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
         Begin Threed.SSPanel pnl_Periodo2 
            Height          =   285
            Left            =   3045
            TabIndex        =   17
            Top             =   315
            Width           =   1405
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Periodo 2"
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
         Begin Threed.SSPanel pnl_Periodo6 
            Height          =   285
            Left            =   8625
            TabIndex        =   18
            Top             =   315
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Periodo 6"
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
         Begin Threed.SSPanel pnl_Periodo3 
            Height          =   285
            Left            =   4440
            TabIndex        =   19
            Top             =   315
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Periodo 3"
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
         Begin Threed.SSPanel pnl_Tipo 
            Height          =   285
            Left            =   60
            TabIndex        =   20
            Top             =   315
            Width           =   1620
            _Version        =   65536
            _ExtentX        =   2857
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Calif. \ Periodo"
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
         Begin Threed.SSPanel pnl_NomEmp 
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   2685
            Width           =   4110
            _Version        =   65536
            _ExtentX        =   7250
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Empresa"
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
         Begin Threed.SSPanel pnl_Periodo4 
            Height          =   285
            Left            =   5835
            TabIndex        =   22
            Top             =   315
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2478
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Periodo 4"
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
         Begin Threed.SSPanel SSPanel31 
            Height          =   285
            Left            =   10020
            TabIndex        =   23
            Top             =   315
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "%"
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
         Begin Threed.SSPanel pnl_MtoTot 
            Height          =   285
            Left            =   11175
            TabIndex        =   24
            Top             =   2685
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total"
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
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Totales ==> S/."
            Height          =   195
            Left            =   75
            TabIndex        =   27
            Top             =   2160
            Width           =   1110
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Resumen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   26
            Top             =   70
            Width           =   795
         End
         Begin VB.Label pnl_Detalle 
            AutoSize        =   -1  'True
            Caption         =   "Detalle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   25
            Top             =   2475
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   28
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
            Height          =   585
            Left            =   660
            TabIndex        =   29
            Top             =   30
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Entidades Técnicas - Datos RCC"
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
            Picture         =   "OpeTra_frm_829.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1155
         Left            =   60
         TabIndex        =   30
         Top             =   1470
         Width           =   12675
         _Version        =   65536
         _ExtentX        =   22357
         _ExtentY        =   2037
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
            TabIndex        =   31
            Top             =   450
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
         Begin Threed.SSPanel pnl_TipDoc 
            Height          =   315
            Left            =   1740
            TabIndex        =   32
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
         Begin Threed.SSPanel pnl_NroDoc 
            Height          =   315
            Left            =   9390
            TabIndex        =   33
            Top             =   120
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
            Left            =   1740
            TabIndex        =   34
            Top             =   780
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
         Begin VB.Label lbl_TipDoc 
            Caption         =   "Tipo Documento:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lbl_NumDoc 
            Caption         =   "Nro. Documento:"
            Height          =   225
            Left            =   7770
            TabIndex        =   37
            Top             =   135
            Width           =   1335
         End
         Begin VB.Label lbl_RazSoc 
            Caption         =   "Razón Social:"
            Height          =   255
            Left            =   150
            TabIndex        =   36
            Top             =   435
            Width           =   1035
         End
         Begin VB.Label lbl_TipEmp 
            Caption         =   "Tipo Empresa:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   780
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   675
         Left            =   60
         TabIndex        =   39
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
         Begin VB.CommandButton cmd_Export 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_829.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Exportar Excel"
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12060
            Picture         =   "OpeTra_frm_829.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_CodSbs As String

Private Sub cmd_Export_Click()
  'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Call fs_GenExc
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt
   Call fs_Inicia
   l_str_CodSbs = moddat_g_str_Codigo
   Call fs_GenRcc
   Call grd_Listad_rcc1_SelChange
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_rcc1_SelChange()
Dim r_str_auxfecha As String
Dim r_int_NroFil   As Integer

    pnl_Detalle.Caption = "Detalle"
    If grd_Listad_rcc1.Rows = 0 Then
       Exit Sub
    End If

    If (grd_Listad_rcc1.Col = 2) Then
        r_str_auxfecha = Trim(pnl_Periodo1.Caption)
    ElseIf (grd_Listad_rcc1.Col = 3) Then
        r_str_auxfecha = Trim(pnl_Periodo2.Caption)
    ElseIf (grd_Listad_rcc1.Col = 4) Then
        r_str_auxfecha = Trim(pnl_Periodo3.Caption)
    ElseIf (grd_Listad_rcc1.Col = 5) Then
        r_str_auxfecha = Trim(pnl_Periodo4.Caption)
    ElseIf (grd_Listad_rcc1.Col = 6) Then
        r_str_auxfecha = Trim(pnl_Periodo5.Caption)
    ElseIf (grd_Listad_rcc1.Col = 7) Then
        r_str_auxfecha = Trim(pnl_Periodo6.Caption)
    End If
    
    pnl_Detalle.Caption = "Detalle Periodo : " & r_str_auxfecha

    For r_int_NroFil = 0 To grd_Listad_rcc2.Rows - 1
       grd_Listad_rcc2.RowHeight(r_int_NroFil) = 0
       If (grd_Listad_rcc2.TextMatrix(r_int_NroFil, 1) = r_str_auxfecha) Then
          grd_Listad_rcc2.RowHeight(r_int_NroFil) = 240
       End If
    Next
    
    'Call gs_RefrescaGrid(grd_Listad_rcc1)
End Sub
Private Sub fs_Inicia()
'Inicializando RCC del Cliente
   grd_Listad_rcc1.ColWidth(0) = 0
   grd_Listad_rcc1.ColWidth(1) = 1580
   grd_Listad_rcc1.ColWidth(2) = 1390
   grd_Listad_rcc1.ColWidth(3) = 1390
   grd_Listad_rcc1.ColWidth(4) = 1390
   grd_Listad_rcc1.ColWidth(5) = 1390
   grd_Listad_rcc1.ColWidth(6) = 1390
   grd_Listad_rcc1.ColWidth(7) = 1390
   grd_Listad_rcc1.ColWidth(8) = 1200
   grd_Listad_rcc1.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Listad_rcc2.ColWidth(0) = 0
   grd_Listad_rcc2.ColWidth(1) = 0
   grd_Listad_rcc2.ColWidth(2) = 4110
   grd_Listad_rcc2.ColWidth(3) = 1600
   grd_Listad_rcc2.ColWidth(4) = 3330
   grd_Listad_rcc2.ColWidth(5) = 0
   grd_Listad_rcc2.ColWidth(6) = 1020
   grd_Listad_rcc2.ColWidth(7) = 1030
   grd_Listad_rcc2.ColWidth(8) = 1030
   grd_Listad_rcc2.ColWidth(9) = 0
   grd_Listad_rcc2.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad_rcc2.ColAlignment(3) = flexAlignCenterCenter
   
   pnl_TipDoc.Caption = moddat_g_str_TipDoc
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri
   
End Sub
Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_FilGrd     As Integer
Dim r_int_FilExl     As Integer
Dim r_int_filCol     As Integer
Dim r_int_fildet     As Integer
Dim r_str_Cadena     As String
Dim r_int_VarAux     As Integer
Dim r_bol_Estado As Boolean
       
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
        r_int_VarAux = 2
       .Range("B" & r_int_VarAux) = "REPORTE CONSOLIDADO CREDITICIO"
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Font.Underline = True
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Font.Bold = True
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Font.Size = 8
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Merge
       .Range("B" & r_int_VarAux).HorizontalAlignment = xlHAlignCenter
            
       .Range("A1:J85").Font.Name = "Arial"
       .Range("A3:J85").Font.Size = 8
       '.Rows("1:100").RowHeight = 11.25
      
       .Columns("J").HorizontalAlignment = xlHAlignCenter
       .Columns("A").ColumnWidth = 5
       .Columns("B").ColumnWidth = 14
       .Columns("C").ColumnWidth = 12
       .Columns("D").ColumnWidth = 14
       .Columns("E").ColumnWidth = 14
       .Columns("F").ColumnWidth = 14
       .Columns("G").ColumnWidth = 13
       .Columns("H").ColumnWidth = 13
       .Columns("I").ColumnWidth = 13
       .Columns("J").ColumnWidth = 12
       'Bordes de las celdas
        r_int_VarAux = 9
       .Range("B" & r_int_VarAux).Borders(xlEdgeLeft).LineStyle = xlContinuous
       .Range("I" & r_int_VarAux).Borders(xlEdgeRight).LineStyle = xlContinuous
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Borders(xlEdgeBottom).LineStyle = xlContinuous
       '.Range("B45:F45").Borders(xlEdgeTop).Weight = xlThin
       
       .Range("B" & r_int_VarAux & ":J" & r_int_VarAux).HorizontalAlignment = xlHAlignCenter
       .Range("B10:B15").HorizontalAlignment = xlHAlignLeft
       .Range("C10:J16").HorizontalAlignment = xlHAlignRight
       .Range("B16").HorizontalAlignment = xlHAlignCenter
       .Cells(4, r_int_VarAux).HorizontalAlignment = xlHAlignCenter
              
       '.Range("H4:I4").Merge
      
       .Cells(4, 2) = "TIPO - NRO. DOCUMENTO:"
       .Cells(5, 2) = "CLIENTE:"
       .Cells(6, 2) = "TIPO EMPRESA:"
       .Cells(4, 4) = Mid(Trim(pnl_TipDoc.Caption), 1, InStr(pnl_TipDoc.Caption, "-") - 1) & "-" & Trim(pnl_NroDoc.Caption)
       .Cells(5, 4) = Trim(pnl_RazSoc.Caption)
       .Cells(6, 4) = Trim(pnl_TipEmp.Caption)
      
        r_int_VarAux = 8
       .Cells(r_int_VarAux, 2) = "RESUMEN"
       .Cells(r_int_VarAux, 2).Font.Bold = True
        r_int_VarAux = 9
       .Cells(r_int_VarAux, 2) = "CLASIF. \ PERIODO"
       .Cells(r_int_VarAux, 3) = Trim(pnl_Periodo1.Caption)
       .Cells(r_int_VarAux, 4) = Trim(pnl_Periodo2.Caption)
       .Cells(r_int_VarAux, 5) = Trim(pnl_Periodo3.Caption)
       .Cells(r_int_VarAux, 6) = Trim(pnl_Periodo4.Caption)
       .Cells(r_int_VarAux, 7) = Trim(pnl_Periodo5.Caption)
       .Cells(r_int_VarAux, 8) = Trim(pnl_Periodo6.Caption)
       .Cells(r_int_VarAux, 9) = "%"
               
       r_int_FilExl = 10
       r_int_filCol = 3
       For r_int_FilGrd = 0 To grd_Listad_rcc1.Rows - 1
           .Cells(r_int_FilExl, 2) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 1)
           .Cells(r_int_FilExl, 3) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 2)
           .Cells(r_int_FilExl, 4) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 3)
           .Cells(r_int_FilExl, 5) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 4)
           .Cells(r_int_FilExl, 6) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 5)
           .Cells(r_int_FilExl, 7) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 6)
           .Cells(r_int_FilExl, 8) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 7)
           .Cells(r_int_FilExl, 9) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 8)
          
           r_int_FilExl = r_int_FilExl + 1
           r_int_filCol = r_int_filCol + 1
       Next
       
       .Range("C11:J16").NumberFormat = "###,###,##0.00"
       r_int_VarAux = 16
       .Cells(r_int_VarAux, 3) = pnl_Total1.Caption
       .Cells(r_int_VarAux, 4) = pnl_Total2.Caption
       .Cells(r_int_VarAux, 5) = pnl_Total3.Caption
       .Cells(r_int_VarAux, 6) = pnl_Total4.Caption
       .Cells(r_int_VarAux, 7) = pnl_Total5.Caption
       .Cells(r_int_VarAux, 8) = pnl_Total6.Caption
       .Cells(r_int_VarAux, 9) = pnl_Total7.Caption
       .Range("B" & r_int_VarAux & ":J" & r_int_VarAux).Font.Bold = True
       .Cells(r_int_VarAux, 2) = "TOTAL"
       r_int_VarAux = 18
       .Cells(r_int_VarAux, 2) = "DETALLE"
       .Cells(r_int_VarAux, 2).Font.Bold = True
       r_int_VarAux = 19
       .Cells(r_int_VarAux, 2) = "NOMBRE EMPRESA"
       .Cells(r_int_VarAux, 4) = "CLASIFICACION"
       .Cells(r_int_VarAux, 5) = "TIPO DEUDA"
       .Cells(r_int_VarAux, 7) = "MONTO (S/.)"
       .Cells(r_int_VarAux, 8) = "MONTO (US$)"
       .Cells(r_int_VarAux, 9) = "TOTAL (S/.)"
      
       'Bordes de las celdas
       .Range("B" & r_int_VarAux).Borders(xlEdgeLeft).LineStyle = xlContinuous
       .Range("I" & r_int_VarAux).Borders(xlEdgeRight).LineStyle = xlContinuous
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Borders(xlEdgeBottom).LineStyle = xlContinuous
       
       .Range("B" & r_int_VarAux & ":C" & r_int_VarAux).Merge
       .Range("E" & r_int_VarAux & ":F" & r_int_VarAux).Merge
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).HorizontalAlignment = xlHAlignCenter
      
       r_str_Cadena = ""
       r_int_FilExl = 20
       r_int_fildet = 0
       r_bol_Estado = False
       
       For r_int_FilGrd = 1 To 6
           Select Case r_int_FilGrd
                  Case 1: r_str_Cadena = pnl_Periodo1.Caption
                  Case 2: r_str_Cadena = pnl_Periodo2.Caption
                  Case 3: r_str_Cadena = pnl_Periodo3.Caption
                  Case 4: r_str_Cadena = pnl_Periodo4.Caption
                  Case 5: r_str_Cadena = pnl_Periodo5.Caption
                  Case 6: r_str_Cadena = pnl_Periodo6.Caption
           End Select
                      
           If (Len(Trim(r_str_Cadena)) > 3) Then
               If (r_bol_Estado = False) Then
                   .Cells(20, 2) = "Detalle del Periodo : " & r_str_Cadena
                   r_int_FilExl = r_int_FilExl + 1
                   .Range("B20:C20").Merge
                   .Range("B20:C20").Font.Bold = True
               Else
                   .Cells(r_int_FilExl, 2) = "Detalle del Periodo : " & r_str_Cadena
                   .Range("B" & r_int_FilExl & ":C" & r_int_FilExl).Merge
                   .Range("B" & r_int_FilExl & ":C" & r_int_FilExl).Font.Bold = True
                   r_int_FilExl = r_int_FilExl + 1
               End If
   
               For r_int_fildet = 0 To grd_Listad_rcc2.Rows - 1
                   If (Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 1)) = r_str_Cadena) Then
                       .Cells(r_int_FilExl, 2) = "     " & Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 2))
                       .Cells(r_int_FilExl, 4) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 3)) '"Clasificación"
                       .Cells(r_int_FilExl, 5) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 4)) '"Tipo Deuda
                       .Cells(r_int_FilExl, 7) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 6))
                       .Cells(r_int_FilExl, 8) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 7))
                       .Cells(r_int_FilExl, 7).NumberFormat = "###,###,##0.00"
                       .Cells(r_int_FilExl, 8).NumberFormat = "###,###,##0.00"
                       .Cells(r_int_FilExl, 9) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 8))
                       .Cells(r_int_FilExl, 9).NumberFormat = "###,###,##0.00"
                       
                       r_int_FilExl = r_int_FilExl + 1
                       r_bol_Estado = True
                   End If
               Next
               r_int_FilExl = r_int_FilExl + 1
           End If
           
           If (r_bol_Estado = False) Then
               r_int_FilExl = 20
           End If
       Next
       
       If (r_bol_Estado = False) Then
           .Cells(20, 2) = ""
       End If
   End With

   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
   
End Sub
Private Sub fs_GenRcc()
Dim r_str_auxfch As String
Dim r_int_Filaux  As Integer
Dim r_str_Cadena  As String
Dim r_int_filCol  As Integer
Dim r_dbl_importe As Double
Dim r_str_PerMes  As String
Dim r_str_PerAno  As String
Dim r_str_perio1  As String
Dim r_str_perio2  As String
Dim r_str_perio3  As String
Dim r_str_perio4  As String
Dim r_str_perio5  As String
Dim r_str_perio6  As String
Dim r_dbl_NumMay As Double
Dim r_dbl_NumFil As Double
      
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT * "
    g_str_Parame = g_str_Parame & "  FROM (SELECT DISTINCT RCCCAB_PERANO, RCCCAB_PERMES "
    g_str_Parame = g_str_Parame & "          FROM CLI_RCCCAB "
    g_str_Parame = g_str_Parame & "         ORDER BY RCCCAB_PERANO DESC, RCCCAB_PERMES DESC) "
    g_str_Parame = g_str_Parame & " WHERE ROWNUM < 2 "
    g_str_Parame = g_str_Parame & " ORDER BY RCCCAB_PERANO DESC "
      
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
   
    r_str_PerMes = g_rst_Princi!RCCCAB_PERMES
    r_str_PerAno = g_rst_Princi!RCCCAB_PERANO
    r_str_Cadena = "01/" & r_str_PerMes & "/" & r_str_PerAno
    
    pnl_Periodo1.Caption = ""
    pnl_Periodo2.Caption = ""
    pnl_Periodo3.Caption = ""
    pnl_Periodo4.Caption = ""
    pnl_Periodo5.Caption = ""
    pnl_Periodo6.Caption = ""
    
    r_str_perio1 = r_str_Cadena
    r_str_perio2 = DateAdd("m", -1, CDate(r_str_Cadena))
    r_str_perio3 = DateAdd("m", -1, CDate(r_str_perio2))
    r_str_perio4 = DateAdd("m", -1, CDate(r_str_perio3))
    r_str_perio5 = DateAdd("m", -1, CDate(r_str_perio4))
    r_str_perio6 = DateAdd("m", -1, CDate(r_str_perio5))
    
    pnl_Periodo1.Caption = Year(r_str_perio6) & "-" & Format(Month(r_str_perio6), "00")
    pnl_Periodo2.Caption = Year(r_str_perio5) & "-" & Format(Month(r_str_perio5), "00")
    pnl_Periodo3.Caption = Year(r_str_perio4) & "-" & Format(Month(r_str_perio4), "00")
    pnl_Periodo4.Caption = Year(r_str_perio3) & "-" & Format(Month(r_str_perio3), "00")
    pnl_Periodo5.Caption = Year(r_str_perio2) & "-" & Format(Month(r_str_perio2), "00")
    pnl_Periodo6.Caption = Year(r_str_perio1) & "-" & Format(Month(r_str_perio1), "00")

    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT RCCCAB_TIPDOC  , RCCCAB_NUMDOC, RCCCAB_PERMES, RCCCAB_PERANO, "
    g_str_Parame = g_str_Parame & "       RCCCAB_CODSBS  , RCCCAB_NUMEMP, RCCCAB_DEUCA0 DEUNOR, RCCCAB_DEUCA1 DEUCPP, "
    g_str_Parame = g_str_Parame & "       RCCCAB_DEUCA2 DEUDEF  , RCCCAB_DEUCA3 DEUDUD, RCCCAB_DEUCA4 DEUPER "
    g_str_Parame = g_str_Parame & "  FROM CLI_RCCCAB "
    g_str_Parame = g_str_Parame & " WHERE ((RCCCAB_PERANO = '" & Left(pnl_Periodo1.Caption, 4) & "' AND RCCCAB_PERMES = '" & Right(pnl_Periodo1.Caption, 2) & "') OR"
    g_str_Parame = g_str_Parame & "        (RCCCAB_PERANO = '" & Left(pnl_Periodo2.Caption, 4) & "' AND RCCCAB_PERMES = '" & Right(pnl_Periodo2.Caption, 2) & "') OR"
    g_str_Parame = g_str_Parame & "        (RCCCAB_PERANO = '" & Left(pnl_Periodo3.Caption, 4) & "' AND RCCCAB_PERMES = '" & Right(pnl_Periodo3.Caption, 2) & "') OR"
    g_str_Parame = g_str_Parame & "        (RCCCAB_PERANO = '" & Left(pnl_Periodo4.Caption, 4) & "' AND RCCCAB_PERMES = '" & Right(pnl_Periodo4.Caption, 2) & "') OR"
    g_str_Parame = g_str_Parame & "        (RCCCAB_PERANO = '" & Left(pnl_Periodo5.Caption, 4) & "' AND RCCCAB_PERMES = '" & Right(pnl_Periodo5.Caption, 2) & "') OR"
    g_str_Parame = g_str_Parame & "        (RCCCAB_PERANO = '" & Left(pnl_Periodo6.Caption, 4) & "' AND RCCCAB_PERMES = '" & Right(pnl_Periodo6.Caption, 2) & "')) "
    g_str_Parame = g_str_Parame & "   AND RCCCAB_CODSBS = '" & l_str_CodSbs & "' "
    g_str_Parame = g_str_Parame & " ORDER BY RCCCAB_PERANO ASC, RCCCAB_PERMES DESC"
   
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
    End If
   
    grd_Listad_rcc1.Rows = 6
    grd_Listad_rcc1.TextMatrix(0, 1) = "Num. Empresas"
    grd_Listad_rcc1.TextMatrix(1, 1) = "D. Normal"
    grd_Listad_rcc1.TextMatrix(2, 1) = "D. CPP"
    grd_Listad_rcc1.TextMatrix(3, 1) = "D. Deficiente"
    grd_Listad_rcc1.TextMatrix(4, 1) = "D. Dudoso"
    grd_Listad_rcc1.TextMatrix(5, 1) = "D. Perdida"
       
    For r_int_Filaux = 0 To grd_Listad_rcc1.Rows - 1
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 2) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 3) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 4) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 5) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 6) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 7) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 8) = "0.00"
    Next

    grd_Listad_rcc2.Rows = 0

    r_int_filCol = 0
    r_int_Filaux = 1
    If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
       g_rst_Princi.MoveFirst
       Do While Not g_rst_Princi.EOF
          r_str_Cadena = Trim(g_rst_Princi!RCCCAB_PERANO) & "-" & Format(Trim(g_rst_Princi!RCCCAB_PERMES), "00")
          If pnl_Periodo1.Caption = r_str_Cadena Then
              r_int_filCol = 2
          ElseIf pnl_Periodo2.Caption = r_str_Cadena Then
              r_int_filCol = 3
          ElseIf pnl_Periodo3.Caption = r_str_Cadena Then
              r_int_filCol = 4
          ElseIf pnl_Periodo4.Caption = r_str_Cadena Then
              r_int_filCol = 5
          ElseIf pnl_Periodo5.Caption = r_str_Cadena Then
              r_int_filCol = 6
          ElseIf pnl_Periodo6.Caption = r_str_Cadena Then
              r_int_filCol = 7
          End If
          
          grd_Listad_rcc1.TextMatrix(0, r_int_filCol) = Trim(g_rst_Princi!RCCCAB_NUMEMP)
          
          If (g_rst_Princi!DEUNOR > 0) Then
              grd_Listad_rcc1.TextMatrix(1, r_int_filCol) = Format(g_rst_Princi!DEUNOR, "###,###,##0.00")
          Else
              grd_Listad_rcc1.TextMatrix(1, r_int_filCol) = "0.00"
          End If
          If (g_rst_Princi!DEUCPP > 0) Then
              grd_Listad_rcc1.TextMatrix(2, r_int_filCol) = Format(g_rst_Princi!DEUCPP, "###,###,##0.00")
          Else
              grd_Listad_rcc1.TextMatrix(2, r_int_filCol) = "0.00"
          End If
          If (g_rst_Princi!DEUDEF > 0) Then
              grd_Listad_rcc1.TextMatrix(3, r_int_filCol) = Format(g_rst_Princi!DEUDEF, "###,###,##0.00")
          Else
              grd_Listad_rcc1.TextMatrix(3, r_int_filCol) = "0.00"
          End If
          If (g_rst_Princi!DEUDUD > 0) Then
              grd_Listad_rcc1.TextMatrix(4, r_int_filCol) = Format(g_rst_Princi!DEUDUD, "###,###,##0.00")
          Else
              grd_Listad_rcc1.TextMatrix(4, r_int_filCol) = "0.00"
          End If
          If (g_rst_Princi!DEUPER > 0) Then
              grd_Listad_rcc1.TextMatrix(5, r_int_filCol) = Format(g_rst_Princi!DEUPER, "###,###,##0.00")
          Else
              grd_Listad_rcc1.TextMatrix(5, r_int_filCol) = "0.00"
          End If
          
          g_rst_Princi.MoveNext
          r_int_Filaux = r_int_Filaux + 1
          DoEvents
       Loop
       Call gs_UbiIniGrid(grd_Listad_rcc1)
       
       'totales del resumen
       For r_int_Filaux = 1 To grd_Listad_rcc1.Rows - 1
           pnl_Total1.Caption = CStr(CDbl(pnl_Total1.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 2)))
           pnl_Total2.Caption = CStr(CDbl(pnl_Total2.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 3)))
           pnl_Total3.Caption = CStr(CDbl(pnl_Total3.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 4)))
           pnl_Total4.Caption = CStr(CDbl(pnl_Total4.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 5)))
           pnl_Total5.Caption = CStr(CDbl(pnl_Total5.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 6)))
           pnl_Total6.Caption = CStr(CDbl(pnl_Total6.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 7)))
       Next
       
      'porcentaje
      r_dbl_NumMay = 0
      r_dbl_NumFil = 0
      r_dbl_importe = 0
      For r_int_Filaux = 1 To grd_Listad_rcc1.Rows - 1
         If grd_Listad_rcc1.TextMatrix(r_int_Filaux, 7) > 0 Then
           r_dbl_importe = (CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 7)) * 100) / CDbl(pnl_Total6.Caption)
           r_dbl_importe = Round(r_dbl_importe, 2)
           If (r_dbl_importe >= r_dbl_NumMay) Then
               r_dbl_NumFil = r_int_Filaux
               r_dbl_NumMay = r_dbl_importe
           End If
         Else
            r_dbl_importe = 0
         End If
           grd_Listad_rcc1.TextMatrix(r_int_Filaux, 8) = Format(r_dbl_importe, "###,###,##0.00")
           pnl_Total7.Caption = CStr(CDbl(pnl_Total7.Caption) + r_dbl_importe)
      Next
       
      'ajuste
       r_dbl_importe = 0
      r_dbl_importe = Round(100 - CDbl(pnl_Total7.Caption), 2)
      If (r_dbl_importe <> CDbl(0)) Then
           grd_Listad_rcc1.TextMatrix(r_dbl_NumFil, 8) = r_dbl_importe + grd_Listad_rcc1.TextMatrix(r_dbl_NumFil, 8)
           pnl_Total7.Caption = CStr(CDbl(pnl_Total7.Caption) + r_dbl_importe)
      End If
       
      pnl_Total1.Caption = Format(CDbl(pnl_Total1.Caption), "###,###,##0.00") & " "
      pnl_Total2.Caption = Format(CDbl(pnl_Total2.Caption), "###,###,##0.00") & " "
      pnl_Total3.Caption = Format(CDbl(pnl_Total3.Caption), "###,###,##0.00") & " "
      pnl_Total4.Caption = Format(CDbl(pnl_Total4.Caption), "###,###,##0.00") & " "
      pnl_Total5.Caption = Format(CDbl(pnl_Total5.Caption), "###,###,##0.00") & " "
      pnl_Total6.Caption = Format(CDbl(pnl_Total6.Caption), "###,###,##0.00") & " "
      pnl_Total7.Caption = Format(CDbl(pnl_Total7.Caption), "###,###,##0.00") & " "
      
'      grd_Listad_rcc1.Row = 0
'      grd_Listad_rcc1.Col = 0

      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT RCCDET_TIPDOC , RCCDET_NUMDOC  , RCCDET_PERMES , RCCDET_PERANO , RCCDET_TIPDEU , RCCDET_DIAATR, "
      g_str_Parame = g_str_Parame & "       RCCDET_CLASIF , RCCDET_MTOSOL, RCCDET_MTODOL, RCCDET_MONDEU, RCCDET_CODEMP "
      g_str_Parame = g_str_Parame & "  FROM CLI_RCCDET "
      g_str_Parame = g_str_Parame & " WHERE ((RCCDET_PERANO = '" & Left(pnl_Periodo1.Caption, 4) & "' AND RCCDET_PERMES = '" & Right(pnl_Periodo1.Caption, 2) & "') OR "
      g_str_Parame = g_str_Parame & "        (RCCDET_PERANO = '" & Left(pnl_Periodo2.Caption, 4) & "' AND RCCDET_PERMES = '" & Right(pnl_Periodo2.Caption, 2) & "') OR "
      g_str_Parame = g_str_Parame & "        (RCCDET_PERANO = '" & Left(pnl_Periodo3.Caption, 4) & "' AND RCCDET_PERMES = '" & Right(pnl_Periodo3.Caption, 2) & "') OR "
      g_str_Parame = g_str_Parame & "        (RCCDET_PERANO = '" & Left(pnl_Periodo4.Caption, 4) & "' AND RCCDET_PERMES = '" & Right(pnl_Periodo4.Caption, 2) & "') OR "
      g_str_Parame = g_str_Parame & "        (RCCDET_PERANO = '" & Left(pnl_Periodo5.Caption, 4) & "' AND RCCDET_PERMES = '" & Right(pnl_Periodo5.Caption, 2) & "') OR "
      g_str_Parame = g_str_Parame & "        (RCCDET_PERANO = '" & Left(pnl_Periodo6.Caption, 4) & "' AND RCCDET_PERMES = '" & Right(pnl_Periodo6.Caption, 2) & "')) "
      g_str_Parame = g_str_Parame & "   AND RCCDET_TIPDOC = '" & IIf(moddat_g_int_TipDoc = 6, 7, moddat_g_int_TipDoc) & "' "
      g_str_Parame = g_str_Parame & "   AND RCCDET_NUMDOC = '" & moddat_g_str_NumDoc & "' "
      
    End If
      
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
   
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
   
    grd_Listad_rcc2.Rows = 0
   
    If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
       g_rst_Princi.MoveFirst
       Do While Not g_rst_Princi.EOF
          'Buscando datos de la Garantía en Registro de Hipotecas
          grd_Listad_rcc2.Rows = grd_Listad_rcc2.Rows + 1
          grd_Listad_rcc2.Row = grd_Listad_rcc2.Rows - 1

          grd_Listad_rcc2.Col = 1
          grd_Listad_rcc2.Text = Trim(g_rst_Princi!RCCDET_PERANO) & "-" & Format(Trim(g_rst_Princi!RCCDET_PERMES), "00")

          grd_Listad_rcc2.Col = 2
          grd_Listad_rcc2.Text = fs_Buscar_NomEmp(Trim(g_rst_Princi!RCCDET_CODEMP))

          grd_Listad_rcc2.Col = 3
          grd_Listad_rcc2.Text = fs_Busca_Clasificacion(Trim(g_rst_Princi!RCCDET_CLASIF))

          grd_Listad_rcc2.Col = 4
          grd_Listad_rcc2.Text = fs_Carga_Creditos(Trim(g_rst_Princi!RCCDET_TIPDEU))

          grd_Listad_rcc2.Col = 5
          If CInt(Trim(g_rst_Princi!RCCDET_MONDEU)) = 1 Or CInt(Trim(g_rst_Princi!RCCDET_MONDEU)) = 3 Then
            grd_Listad_rcc2.Text = "SOLES"
          ElseIf CInt(Trim(g_rst_Princi!RCCDET_MONDEU)) = 2 Then
            grd_Listad_rcc2.Text = "DOLARES AMERICANOS"
          End If
          'grd_Listad_rcc2.Text = Trim(g_rst_Princi!RCCDET_MONDEU) 'RCCDET_TIPMON

          grd_Listad_rcc2.Col = 6
          grd_Listad_rcc2.Text = Format(g_rst_Princi!RCCDET_MTOSOL, "###,###,##0.00")

          grd_Listad_rcc2.Col = 7
          grd_Listad_rcc2.Text = Format(g_rst_Princi!RCCDET_MTODOL, "###,###,##0.00")
          
          r_dbl_importe = CDbl(IIf(IsNull(g_rst_Princi!RCCDET_MTOSOL) = True, "0.00", g_rst_Princi!RCCDET_MTOSOL)) + _
                          CDbl(IIf(IsNull(g_rst_Princi!RCCDET_MTODOL) = True, "0.00", g_rst_Princi!RCCDET_MTODOL))
          grd_Listad_rcc2.Col = 8
          grd_Listad_rcc2.Text = Format(r_dbl_importe, "###,###,##0.00")
          
          grd_Listad_rcc2.Col = 9
          grd_Listad_rcc2.Text = g_rst_Princi!RCCDET_DIAATR
          
          g_rst_Princi.MoveNext
          DoEvents
       Loop

'       grd_Listad_rcc2.Row = 0
'       grd_Listad_rcc2.Col = 0
       
       Call gs_UbiIniGrid(grd_Listad_rcc2)
    End If
    
                   
'    grd_Listad_rcc2.Col = 2
'    grd_Listad_rcc2.Sort = 7
                   
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
End Sub
Private Function fs_Busca_Clasificacion(ByVal p_CodCla As String) As String
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TIPCLA_DESCRI "
   g_str_Parame = g_str_Parame & "  FROM CTB_TIPCLA "
   g_str_Parame = g_str_Parame & " WHERE TIPCLA_TIPCRE = 13 "
   g_str_Parame = g_str_Parame & "   AND TIPCLA_CODIGO = " & p_CodCla & ""
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
'      If Trim(g_rst_Listas!TIPCLA_DESCRI) = "CON PROBLEMA POTENCIAL" Then
'         fs_Busca_Clasificacion = "CPP"
'      Else
         fs_Busca_Clasificacion = Trim(g_rst_Listas!TIPCLA_DESCRI)
'      End If
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function
Private Function fs_Carga_Creditos(ByVal p_TipDeu As Integer) As String
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TIPCRE_DESCRI "
   g_str_Parame = g_str_Parame & "  FROM CTB_TIPCRE "
   g_str_Parame = g_str_Parame & " WHERE TIPCRE_CODIGO = " & p_TipDeu & ""
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      fs_Carga_Creditos = Trim(g_rst_Listas!TIPCRE_DESCRI)
   End If
    
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function
Private Function fs_Buscar_NomEmp(ByVal p_CodEmp As Integer) As String
   fs_Buscar_NomEmp = ""
   
   g_str_Parame = "SELECT * FROM CTB_EMPSUP WHERE EMPSUP_CODIGO = " & p_CodEmp & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      fs_Buscar_NomEmp = Trim(g_rst_Listas!EMPSUP_NOMBRE)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub pnl_MtoDol_Click()
  If Len(Trim(pnl_MtoDol.Tag)) = 0 Or pnl_MtoDol.Tag = "D" Then
      pnl_MtoDol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad_rcc2, 7, "N")
  Else
      pnl_MtoDol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad_rcc2, 7, "N-")
  End If
End Sub

Private Sub pnl_MtoSol_Click()
  If Len(Trim(pnl_MtoSol.Tag)) = 0 Or pnl_MtoSol.Tag = "D" Then
      pnl_MtoSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad_rcc2, 6, "N")
  Else
      pnl_MtoSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad_rcc2, 6, "N-")
  End If
End Sub

Private Sub pnl_MtoTot_Click()
  If Len(Trim(pnl_MtoTot.Tag)) = 0 Or pnl_MtoTot.Tag = "D" Then
      pnl_MtoTot.Tag = "A"
      Call gs_SorteaGrid(grd_Listad_rcc2, 8, "N")
  Else
      pnl_MtoTot.Tag = "D"
      Call gs_SorteaGrid(grd_Listad_rcc2, 8, "N-")
  End If
End Sub

Private Sub pnl_NomEmp_Click()
  If Len(Trim(pnl_NomEmp.Tag)) = 0 Or pnl_NomEmp.Tag = "D" Then
      pnl_NomEmp.Tag = "A"
      Call gs_SorteaGrid(grd_Listad_rcc2, 2, "C")
  Else
      pnl_NomEmp.Tag = "D"
      Call gs_SorteaGrid(grd_Listad_rcc2, 2, "C-")
  End If
End Sub

Private Sub pnl_TipCla_Click()
  If Len(Trim(pnl_TipCla.Tag)) = 0 Or pnl_TipCla.Tag = "D" Then
      pnl_TipCla.Tag = "A"
      Call gs_SorteaGrid(grd_Listad_rcc2, 3, "C")
  Else
      pnl_TipCla.Tag = "D"
      Call gs_SorteaGrid(grd_Listad_rcc2, 3, "C-")
  End If
End Sub

Private Sub pnl_TipDeu_Click()
  If Len(Trim(pnl_TipDeu.Tag)) = 0 Or pnl_TipDeu.Tag = "D" Then
      pnl_TipDeu.Tag = "A"
      Call gs_SorteaGrid(grd_Listad_rcc2, 4, "C")
  Else
      pnl_TipDeu.Tag = "D"
      Call gs_SorteaGrid(grd_Listad_rcc2, 4, "C-")
  End If
End Sub
