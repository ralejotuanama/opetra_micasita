VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ges_TecPro_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   Icon            =   "OpeTra_frm_827.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   10550
      Left            =   30
      TabIndex        =   27
      Top             =   0
      Width           =   12000
      _Version        =   65536
      _ExtentX        =   21167
      _ExtentY        =   18609
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   495
         Left            =   30
         TabIndex        =   76
         Top             =   6780
         Width           =   11895
         _Version        =   65536
         _ExtentX        =   20981
         _ExtentY        =   873
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
         BevelOuter      =   1
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   120
            Width           =   3975
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Situación:"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   180
            Width           =   705
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1935
         Left            =   30
         TabIndex        =   39
         Top             =   4800
         Width           =   11925
         _Version        =   65536
         _ExtentX        =   21034
         _ExtentY        =   3413
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
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   18
            Text            =   "cmb_DstDir"
            Top             =   1560
            Width           =   3975
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   16
            Text            =   "cmb_DptDir"
            Top             =   1200
            Width           =   3975
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8430
            TabIndex        =   17
            Text            =   "cmb_PrvDir"
            Top             =   1200
            Width           =   3345
         End
         Begin VB.ComboBox cmb_TdoRep 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox txt_TelRep 
            Height          =   315
            Left            =   8430
            MaxLength       =   25
            TabIndex        =   19
            Top             =   1560
            Width           =   3315
         End
         Begin VB.TextBox txt_DirRep 
            Height          =   315
            Left            =   2010
            MaxLength       =   120
            TabIndex        =   15
            Top             =   840
            Width           =   9765
         End
         Begin VB.TextBox txt_NdoRep 
            Height          =   315
            Left            =   8430
            MaxLength       =   12
            TabIndex        =   14
            Top             =   510
            Width           =   3345
         End
         Begin VB.TextBox txt_RepLeg 
            Height          =   315
            Left            =   2010
            MaxLength       =   100
            TabIndex        =   12
            Top             =   120
            Width           =   9765
         End
         Begin VB.Label Label21 
            Caption         =   "Distrito:"
            Height          =   225
            Left            =   120
            TabIndex        =   53
            Top             =   1605
            Width           =   885
         End
         Begin VB.Label Label19 
            Caption         =   "Departamento:"
            Height          =   225
            Left            =   120
            TabIndex        =   52
            Top             =   1245
            Width           =   1065
         End
         Begin VB.Label Label20 
            Caption         =   "Provincia:"
            Height          =   255
            Left            =   6870
            TabIndex        =   51
            Top             =   1230
            Width           =   825
         End
         Begin VB.Label Label13 
            Caption         =   "Tipo Documento:"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   510
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Teléfono:"
            Height          =   195
            Left            =   6840
            TabIndex        =   43
            Top             =   1620
            Width           =   795
         End
         Begin VB.Label Label9 
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   870
            Width           =   1545
         End
         Begin VB.Label Label8 
            Caption         =   "Nro. Documento:"
            Height          =   195
            Left            =   6870
            TabIndex        =   41
            Top             =   570
            Width           =   1305
         End
         Begin VB.Label Label7 
            Caption         =   "Representante Legal"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   150
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1665
         Left            =   30
         TabIndex        =   37
         Top             =   7860
         Width           =   11925
         _Version        =   65536
         _ExtentX        =   21034
         _ExtentY        =   2937
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
         Begin Threed.SSPanel pnl_LinAsig_Ind 
            Height          =   315
            Left            =   2010
            TabIndex        =   20
            Top             =   480
            Width           =   2625
            _Version        =   65536
            _ExtentX        =   4630
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_LinAsig_Dir 
            Height          =   315
            Left            =   8430
            TabIndex        =   55
            Top             =   480
            Width           =   2625
            _Version        =   65536
            _ExtentX        =   4630
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_LinRev_Dir 
            Height          =   315
            Left            =   8640
            TabIndex        =   59
            Top             =   840
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_LinNRe_Dir 
            Height          =   315
            Left            =   8640
            TabIndex        =   60
            Top             =   1200
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_LinRev_Ind 
            Height          =   315
            Left            =   2220
            TabIndex        =   72
            Top             =   840
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_LinNRe_Ind 
            Height          =   315
            Left            =   2220
            TabIndex        =   73
            Top             =   1200
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Revolvente:"
            Height          =   195
            Left            =   240
            TabIndex        =   75
            Top             =   900
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "No Revolvente:"
            Height          =   195
            Left            =   240
            TabIndex        =   74
            Top             =   1260
            Width           =   1125
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "No Revolvente:"
            Height          =   195
            Left            =   6960
            TabIndex        =   62
            Top             =   1260
            Width           =   1125
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Revolvente:"
            Height          =   195
            Left            =   6960
            TabIndex        =   61
            Top             =   900
            Width           =   870
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Créditos Directos"
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
            Left            =   6870
            TabIndex        =   58
            Top             =   120
            Width           =   1470
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Créditos Indirectos"
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
            Left            =   120
            TabIndex        =   57
            Top             =   120
            Width           =   1605
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Línea Asignada:"
            Height          =   195
            Left            =   6870
            TabIndex        =   56
            Top             =   540
            Width           =   1170
         End
         Begin VB.Label lbl_LinAsig 
            AutoSize        =   -1  'True
            Caption         =   "Línea Asignada:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   540
            Width           =   1170
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   3255
         Left            =   30
         TabIndex        =   28
         Top             =   1500
         Width           =   11925
         _Version        =   65536
         _ExtentX        =   21034
         _ExtentY        =   5741
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
         Begin VB.TextBox txt_CodSbs 
            Height          =   315
            Left            =   8460
            MaxLength       =   10
            TabIndex        =   4
            Top             =   930
            Width           =   3315
         End
         Begin VB.TextBox txt_SegNom 
            Height          =   315
            Left            =   8460
            MaxLength       =   40
            TabIndex        =   9
            Top             =   2010
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeCas 
            Height          =   315
            Left            =   8460
            MaxLength       =   40
            TabIndex        =   7
            Top             =   1650
            Width           =   3315
         End
         Begin VB.TextBox txt_PriNom 
            Height          =   315
            Left            =   2010
            MaxLength       =   40
            TabIndex        =   8
            Top             =   2010
            Width           =   3975
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   2010
            MaxLength       =   40
            TabIndex        =   6
            Top             =   1650
            Width           =   3975
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   2010
            MaxLength       =   40
            TabIndex        =   5
            Top             =   1290
            Width           =   3975
         End
         Begin VB.TextBox txt_NomCor 
            Height          =   315
            Left            =   2010
            MaxLength       =   20
            TabIndex        =   3
            Top             =   930
            Width           =   3975
         End
         Begin VB.ComboBox cmb_CodCiu 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2760
            Width           =   9765
         End
         Begin VB.ComboBox cmb_TipEmp 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2370
            Width           =   3975
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   8460
            MaxLength       =   11
            TabIndex        =   1
            Top             =   180
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   3975
         End
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   315
            Left            =   2010
            TabIndex        =   2
            Top             =   570
            Width           =   9765
            _Version        =   65536
            _ExtentX        =   17224
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
         Begin VB.Label Label22 
            Caption         =   "Código SBS"
            Height          =   255
            Left            =   6870
            TabIndex        =   54
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Segundo Nombre:"
            Height          =   255
            Left            =   6870
            TabIndex        =   50
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Apellido Casado"
            Height          =   255
            Left            =   6870
            TabIndex        =   49
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Primer Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Apellido Materno:"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "Apellido Paterno:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Nombre Corto:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "CIIU"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   2790
            Width           =   1335
         End
         Begin VB.Label lbl_TipEmp 
            Caption         =   "Tipo Empresa:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label lbl_RazSoc 
            Caption         =   "Razón Social:"
            Height          =   255
            Left            =   150
            TabIndex        =   31
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lbl_NumDoc 
            Caption         =   "Nro. Documento:"
            Height          =   225
            Left            =   6870
            TabIndex        =   30
            Top             =   225
            Width           =   1335
         End
         Begin VB.Label lbl_TipDoc 
            Caption         =   "Tipo Documento:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   210
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   33
         Top             =   60
         Width           =   11925
         _Version        =   65536
         _ExtentX        =   21034
         _ExtentY        =   1191
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
            TabIndex        =   34
            Top             =   30
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Entidades Técnicas - Registro"
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
            Picture         =   "OpeTra_frm_827.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   30
         TabIndex        =   35
         Top             =   780
         Width           =   11925
         _Version        =   65536
         _ExtentX        =   21034
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
         Begin VB.CommandButton cmd_EteHis 
            Height          =   585
            Left            =   1740
            Picture         =   "OpeTra_frm_827.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Histórico"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11310
            Picture         =   "OpeTra_frm_827.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatAcc 
            Enabled         =   0   'False
            Height          =   585
            Left            =   1170
            Picture         =   "OpeTra_frm_827.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Accionistas"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VtaPat 
            Height          =   585
            Left            =   600
            Picture         =   "OpeTra_frm_827.frx":0FDC
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Ventas y Patrimonio"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatRCC 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_827.frx":12E6
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Datos del RCC"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   10740
            Picture         =   "OpeTra_frm_827.frx":1728
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   900
         Left            =   30
         TabIndex        =   63
         Top             =   9570
         Width           =   11925
         _Version        =   65536
         _ExtentX        =   21034
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
         Begin Threed.SSPanel pnl_PorRet 
            Height          =   315
            Left            =   2010
            TabIndex        =   68
            Top             =   120
            Width           =   2625
            _Version        =   65536
            _ExtentX        =   4630
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_FecApr 
            Height          =   315
            Left            =   2010
            TabIndex        =   69
            Top             =   480
            Width           =   2625
            _Version        =   65536
            _ExtentX        =   4630
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
         Begin Threed.SSPanel pnl_PorTEA 
            Height          =   315
            Left            =   8430
            TabIndex        =   70
            Top             =   120
            Width           =   2625
            _Version        =   65536
            _ExtentX        =   4630
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_FecVct 
            Height          =   315
            Left            =   8430
            TabIndex        =   71
            Top             =   480
            Width           =   2625
            _Version        =   65536
            _ExtentX        =   4630
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "F. Aprobación"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   540
            Width           =   990
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "% TEA"
            Height          =   195
            Left            =   6870
            TabIndex        =   66
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "F. Vencimiento:"
            Height          =   195
            Left            =   6870
            TabIndex        =   65
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "FMV - % Retención"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   180
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   495
         Left            =   30
         TabIndex        =   79
         Top             =   7320
         Width           =   11895
         _Version        =   65536
         _ExtentX        =   20981
         _ExtentY        =   873
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
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_LinTot 
            Height          =   315
            Left            =   2010
            TabIndex        =   80
            Top             =   120
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ExpLin 
            Height          =   315
            Left            =   8430
            TabIndex        =   83
            Top             =   120
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label Label28 
            Caption         =   "Línea Total Asignada:"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   150
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Exp. Máx. Línea:"
            Height          =   255
            Left            =   6870
            TabIndex        =   81
            Top             =   150
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_TipDoc        As Integer
Dim l_str_NumDoc        As String
Dim l_dbl_LinAsi        As Double
Dim l_dbl_PorRet        As Double
Dim l_str_FecVct        As String

Dim l_arr_PryVin()      As moddat_tpo_Genera
Dim l_str_DptDir        As String
Dim l_str_PrvDir        As String
Dim l_str_DstDir        As String
Dim l_int_FlgCmb        As Integer


Private Sub cmb_CodCiu_Click()
    Call gs_SetFocus(txt_RepLeg) 'ipp_LinAsig
End Sub

Private Sub cmb_CodCiu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodCiu_Click
   End If
End Sub

Private Sub cmb_DptDir_Change()
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_Click()
   If cmb_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_DptDir_GotFocus()
   Call SendMessage(cmb_DptDir.hwnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir, l_str_DptDir)
      l_int_FlgCmb = True
      
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      If cmb_DptDir.ListIndex > -1 Then
         l_str_DptDir = ""
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir)
   End If
End Sub

Private Sub cmb_DptDir_LostFocus()
   Call SendMessage(cmb_DptDir.hwnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_DstDir_Change()
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_Click()
   If cmb_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_TelRep)
      End If
   End If
End Sub

Private Sub cmb_DstDir_GotFocus()
   Call SendMessage(cmb_DstDir.hwnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir, l_str_DstDir)
      l_int_FlgCmb = True
      
      If cmb_DstDir.ListIndex > -1 Then
         l_str_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_TelRep)
   End If
End Sub

Private Sub cmb_DstDir_LostFocus()
   Call SendMessage(cmb_DstDir.hwnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_PrvDir_Change()
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_Click()
   If cmb_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstDir.Clear
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_GotFocus()
   Call SendMessage(cmb_PrvDir.hwnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir, l_str_PrvDir)
      l_int_FlgCmb = True
      
      cmb_DstDir.Clear
      If cmb_PrvDir.ListIndex > -1 Then
         l_str_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir)
   End If
End Sub

Private Sub cmb_PrvDir_LostFocus()
   Call SendMessage(cmb_PrvDir.hwnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   Call cmb_Situac_Click
End Sub

Private Sub cmb_TdoRep_Click()
  Call gs_SetFocus(txt_NdoRep)
End Sub

Private Sub cmb_TdoRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NdoRep)
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   Call gs_SetFocus(txt_NumDoc)
   If cmb_TipDoc.ListIndex <> -1 Then
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 1 Then
         txt_NomCor.Enabled = True
         txt_ApePat.Enabled = False
         txt_ApeMat.Enabled = False
         txt_ApeCas.Enabled = False
         txt_PriNom.Enabled = False
         txt_SegNom.Enabled = False
      Else
         txt_NomCor.Enabled = False
         txt_ApePat.Enabled = True
         txt_ApeMat.Enabled = True
         txt_ApeCas.Enabled = True
         txt_PriNom.Enabled = True
         txt_SegNom.Enabled = True
      End If
   End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
   End If
End Sub

Private Sub cmb_TipEmp_Click()
   Call gs_SetFocus(cmb_CodCiu)
End Sub

Private Sub cmb_TipEmp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodCiu)
   End If
End Sub

Private Sub cmd_DatRCC_Click()
   If moddat_g_str_Codigo <> "" Then
      frm_Ges_TecPro_07.Show 1
   Else
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub
Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DISTINCT RCCCAB_CODSBS FROM CLI_RCCCAB  "
   g_str_Parame = g_str_Parame & " WHERE RCCCAB_TIPDOC = '" & CStr(p_TipDoc) & "'"
   g_str_Parame = g_str_Parame & "   AND RCCCAB_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_g_str_Codigo = IIf(IsNull(g_rst_Princi!RCCCAB_CODSBS), "", Trim(g_rst_Princi!RCCCAB_CODSBS))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_EteHis_Click()
   If modgen_g_int_TipUsu = "18200" Or modgen_g_int_TipUsu = "18970" Or modgen_g_int_TipUsu = "18600" Then
      frm_Ges_TecPro_10.Show 1
   Else
      MsgBox "No tiene acceso a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   If txt_NumDoc.Text = Empty Then
      MsgBox "Debe ingresar Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   If Len(Trim(pnl_RazSoc.Caption)) = 0 Then
       MsgBox "Debe ingresar un documento válido", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(txt_NumDoc)
       Exit Sub
   End If
   If Len(Trim(txt_CodSbs.Text)) = 0 Then
      MsgBox "Debe ingresar el Código SBS.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_CodSbs)
      Exit Sub
   End If
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
      If Len(Trim(txt_ApePat.Text)) = 0 Then
          MsgBox "Debe ingresar Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(txt_ApePat)
          Exit Sub
      End If
      If Len(Trim(txt_ApeMat.Text)) = 0 Then
          MsgBox "Debe ingresar Apellido Materno.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(txt_ApeMat)
          Exit Sub
      End If
      If Len(Trim(txt_PriNom.Text)) = 0 Then
          MsgBox "Debe ingresar Primer Nombre.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(txt_PriNom)
          Exit Sub
      End If
'      If Len(Trim(txt_SegNom.Text)) = 0 Then
'          MsgBox "Debe ingresar Segundo Nombre.", vbExclamation, modgen_g_str_NomPlt
'          Call gs_SetFocus(txt_SegNom)
'          Exit Sub
'      End If
   Else
      If Len(Trim(txt_NomCor.Text)) = 0 Then
          MsgBox "Debe ingresar Nombre Corto de la ET.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(txt_NomCor)
          Exit Sub
      End If
   End If
   
   If cmb_TipEmp.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipEmp)
      Exit Sub
   End If
   
   If cmb_CodCiu.ListIndex = -1 Then
      MsgBox "Debe seleccionar Actividad Económica.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodCiu)
      Exit Sub
   End If
   
'   If cmb_DptDir.ListIndex = -1 Then
'         MsgBox "Debe seleccionar el Departamento.", vbExclamation, modgen_g_str_NomPlt
'         Call gs_SetFocus(cmb_DptDir)
'         Exit Sub
'   End If
'   If cmb_PrvDir.ListIndex = -1 Then
'      MsgBox "Debe seleccionar la Provincia.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(cmb_PrvDir)
'      Exit Sub
'   End If
'   If cmb_DstDir.ListIndex = -1 Then
'      MsgBox "Debe seleccionar el Distrito.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(cmb_DstDir)
'      Exit Sub
'   End If
   
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar Situación de la Entidad", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      If fs_Validar_EntTec = True Then
         MsgBox "Ya se encuentra registrada la Entidad Técnica.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipDoc)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      Call moddat_gs_FecSis
   
      'Grabando Información de Entidad Técnica
      g_str_Parame = "USP_TPR_MAEETE ("
      g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(txt_NumDoc.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(txt_NomCor.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(txt_CodSbs.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(txt_ApePat.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(txt_ApeMat.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(txt_ApeCas.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(txt_PriNom.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(txt_SegNom.Text) & "', "
      
      g_str_Parame = g_str_Parame & CStr(cmb_TipEmp.ItemData(cmb_TipEmp.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)) & ", "
      
      If Len(txt_RepLeg.Text) = 0 Then
         g_str_Parame = g_str_Parame & "'', "
      Else
         g_str_Parame = g_str_Parame & "'" & CStr(txt_RepLeg.Text) & "', "
      End If
      If cmb_TdoRep.ListIndex <> -1 Then
         g_str_Parame = g_str_Parame & CStr(cmb_TdoRep.ItemData(cmb_TdoRep.ListIndex)) & ", "
      Else
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      If Len(txt_NdoRep.Text) = 0 Then
         g_str_Parame = g_str_Parame & "'', "
      Else
         g_str_Parame = g_str_Parame & "'" & CStr(txt_NdoRep.Text) & "', "
      End If
      If Len(txt_DirRep.Text) = 0 Then
         g_str_Parame = g_str_Parame & "'', "
      Else
         g_str_Parame = g_str_Parame & "'" & CStr(txt_DirRep.Text) & "', "
      End If
      If cmb_DptDir.ListIndex = -1 And cmb_PrvDir.ListIndex = -1 And cmb_DstDir.ListIndex = -1 Then
         g_str_Parame = g_str_Parame & "'', "
      Else
         g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
      End If
      If Len(txt_TelRep.Text) = 0 Then
         g_str_Parame = g_str_Parame & "'', "
      Else
         g_str_Parame = g_str_Parame & "'" & CStr(txt_TelRep.Text) & "', "
      End If
      
      g_str_Parame = g_str_Parame & CDbl(pnl_LinAsig_Ind.Caption) & ", "
      g_str_Parame = g_str_Parame & CDbl(0) & ", "                                        'ipp_AdmFlj.Text
      g_str_Parame = g_str_Parame & CDbl(0) & ", "                                        'ipp_ImpHip.Text
      g_str_Parame = g_str_Parame & CDbl(0) & ", "                                        'ipp_ImpLiq.Text

      g_str_Parame = g_str_Parame & CStr(pnl_PorRet.Caption) & ", "
      g_str_Parame = g_str_Parame & "'" & Format(pnl_FecApr.Caption, "yyyymmdd") & "', "
      g_str_Parame = g_str_Parame & "'" & Format(pnl_FecVct.Caption, "yyyymmdd") & "', "
      g_str_Parame = g_str_Parame & CStr(pnl_PorTEA.Caption) & ", "
      g_str_Parame = g_str_Parame & CDbl(pnl_LinAsig_Dir.Caption) & ", "
      g_str_Parame = g_str_Parame & cmb_Situac.ItemData(cmb_Situac.ListIndex) & ", "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   If moddat_g_int_FlgAct_2 = 1 Then
      Call fs_Activa(True)
   End If
   MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_con_PltPar
   
   Call fs_Limpia
   frm_Ges_TecPro_01.fs_Buscar
   frm_Ges_TecPro_01.fs_Activa (True)
   Unload Me
End Sub

Private Function fs_Validar_EntTec() As Boolean
   fs_Validar_EntTec = False

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT COUNT(MAEETE_NUMDOC) CONTADOR "
   g_str_Parame = g_str_Parame & "     FROM TPR_MAEETE"
   g_str_Parame = g_str_Parame & "    WHERE MAEETE_TIPDOC = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " "
   g_str_Parame = g_str_Parame & "      AND MAEETE_NUMDOC = '" & CStr(txt_NumDoc.Text) & "'"
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If g_rst_GenAux!CONTADOR > 0 Then
         fs_Validar_EntTec = True
      End If
   End If
End Function

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VtaPat_Click()
   frm_Ges_TecPro_08.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_Activa(False)
   Call fs_Limpia
   
   If moddat_g_int_FlgGrb = 2 Then
      Call fs_Buscar
      moddat_g_str_Codigo = "" 'l_str_CodSbs = ""
      Call fs_DatCli(IIf(moddat_g_int_TipDoc = 6, 7, moddat_g_int_TipDoc), moddat_g_str_NumDoc)
   End If
      
   Call gs_SetFocus(cmb_TipDoc)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Tipo de Documento
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "118") '203
   Call moddat_gs_Carga_LisIte_Combo(cmb_TdoRep, 1, "118") '203
   
   'Tipo de Empresa
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipEmp, 1, "526")
   
   'Código de Actividad Económica
   Call moddat_gs_Carga_LisIte_Combo(cmb_CodCiu, 1, "102")
   
   'Departamento
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   
   'Situación
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
End Sub

Public Sub fs_Buscar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAEETE_TIPDOC    , MAEETE_NUMDOC, MAEETE_TIPEMP, MAEETE_LINASI_IND , MAEETE_LINASI_DIR  , MAEPRV_RAZSOC   , MAEETE_PORRET, MAEETE_PORTEA    , MAEETE_FECVCT     , MAEETE_CODCIU    , "
   g_str_Parame = g_str_Parame & "        MAEETE_NOMCOR    , MAEETE_APEPAT, MAEETE_APEMAT, MAEETE_APECAS     , MAEETE_PRINOM      , MAEETE_SEGNOM   , MAEETE_NOMREP, MAEETE_TDOREP    , MAEETE_NDOREP     , MAEETE_DIRREP    , "
   g_str_Parame = g_str_Parame & "        MAEETE_TELREP    , MAEETE_ADMFLJ, MAEETE_IMPHIP, MAEETE_IMPLIQ     , MAEETE_FECAPR      , MAEETE_UBIGEO   , MAEETE_CODSBS, MAEETE_LINREV_DIR, MAEETE_LINNRE_DIR , MAEETE_LINREV_IND, "
   g_str_Parame = g_str_Parame & "        MAEETE_LINNRE_IND, MAEETE_LINEXP, MAEETE_LINASI, MAEETE_SITUAC "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAEETE INNER JOIN CNTBL_MAEPRV ON MAEETE_TIPDOC = MAEPRV_TIPDOC AND MAEETE_NUMDOC = MAEPRV_NUMDOC "
   g_str_Parame = g_str_Parame & "  WHERE MAEETE_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND MAEETE_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
       
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   cmb_TipDoc.Text = moddat_gf_Consulta_ParDes("118", g_rst_Princi!MAEETE_TIPDOC)
   txt_NumDoc.Text = g_rst_Princi!MAEETE_NUMDOC
   pnl_RazSoc.Caption = g_rst_Princi!MAEPRV_RAZSOC
   txt_NomCor.Text = IIf(IsNull(g_rst_Princi!MAEETE_NOMCOR), "", Trim(g_rst_Princi!MAEETE_NOMCOR))
   txt_ApePat.Text = IIf(IsNull(g_rst_Princi!MAEETE_APEPAT), "", Trim(g_rst_Princi!MAEETE_APEPAT))
   txt_ApeMat.Text = IIf(IsNull(g_rst_Princi!MAEETE_APEMAT), "", Trim(g_rst_Princi!MAEETE_APEMAT))
   txt_ApeCas.Text = IIf(IsNull(g_rst_Princi!MAEETE_APECAS), "", Trim(g_rst_Princi!MAEETE_APECAS))
   txt_PriNom.Text = IIf(IsNull(g_rst_Princi!MAEETE_PRINOM), "", Trim(g_rst_Princi!MAEETE_PRINOM))
   txt_SegNom.Text = IIf(IsNull(g_rst_Princi!MAEETE_SEGNOM), "", Trim(g_rst_Princi!MAEETE_SEGNOM))
   
   If Not IsNull(g_rst_Princi!MAEETE_CODSBS) Then
      txt_CodSbs.Text = Trim(g_rst_Princi!MAEETE_CODSBS)
   End If
  
   cmb_TipEmp.Text = moddat_gf_Consulta_ParDes("526", g_rst_Princi!MAEETE_TIPEMP)
   cmb_CodCiu.Text = moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!MAEETE_CODCIU))
   
   txt_RepLeg.Text = IIf(IsNull(g_rst_Princi!MAEETE_NOMREP), "", Trim(g_rst_Princi!MAEETE_NOMREP))
   If Not IsNull(g_rst_Princi!MAEETE_TDOREP) Then
      cmb_TdoRep.Text = moddat_gf_Consulta_ParDes("118", g_rst_Princi!MAEETE_TDOREP)
   End If
   
   txt_NdoRep.Text = IIf(IsNull(g_rst_Princi!MAEETE_NDOREP), "", Trim(g_rst_Princi!MAEETE_NDOREP))
   txt_DirRep.Text = IIf(IsNull(g_rst_Princi!MAEETE_DIRREP), "", Trim(g_rst_Princi!MAEETE_DIRREP))
   txt_TelRep.Text = IIf(IsNull(g_rst_Princi!MAEETE_TELREP), "", Trim(g_rst_Princi!MAEETE_TELREP))
   
   cmb_Situac.Text = moddat_gf_Consulta_ParDes("013", g_rst_Princi!MAEETE_SITUAC)
      
   If Not IsNull(g_rst_Princi!MAEETE_LINASI_IND) Then
      pnl_LinAsig_Ind.Caption = Format(CStr(g_rst_Princi!MAEETE_LINASI_IND), "###,##0.00") & "  "
   End If
   If Not IsNull(g_rst_Princi!MAEETE_LINREV_IND) Then
      pnl_LinRev_Ind.Caption = Format(CStr(g_rst_Princi!MAEETE_LINREV_IND), "###,###,###,##0.00") & "  "
   End If
   If Not IsNull(g_rst_Princi!MAEETE_LINNRE_IND) Then
      pnl_LinNRe_Ind.Caption = Format(CStr(g_rst_Princi!MAEETE_LINNRE_IND), "###,###,###,##0.00") & "  "
   End If
   pnl_LinAsig_Dir.Caption = Format(CStr(g_rst_Princi!MAEETE_LINASI_DIR), "###,##0.00") & "  "
   If Not IsNull(g_rst_Princi!MAEETE_LINREV_DIR) Then
      pnl_LinRev_Dir.Caption = Format(CStr(g_rst_Princi!MAEETE_LINREV_DIR), "###,###,###,##0.00") & "  "
   End If
   If Not IsNull(g_rst_Princi!MAEETE_LINREV_DIR) Then
      pnl_LinNRe_Dir.Caption = Format(CStr(g_rst_Princi!MAEETE_LINNRE_DIR), "###,###,###,##0.00") & "  "
   End If
   If Not IsNull(g_rst_Princi!MAEETE_LINASI) Then
      pnl_LinTot.Caption = Format(CStr(g_rst_Princi!MAEETE_LINASI), "###,###,###,##0.00") & "  "
   End If
      If Not IsNull(g_rst_Princi!MAEETE_LINEXP) Then
      pnl_ExpLin.Caption = Format(CStr(g_rst_Princi!MAEETE_LINEXP), "###,###,###,##0.00") & "  "
   End If
   If Not IsNull(g_rst_Princi!MAEETE_PORRET) Then
      pnl_PorRet.Caption = Format(IIf(IsNull(g_rst_Princi!MAEETE_PORRET), 0, g_rst_Princi!MAEETE_PORRET), "0.00") & "  "
   End If
   If Not IsNull(g_rst_Princi!MAEETE_PORTEA) Then
      pnl_PorTEA.Caption = Format(IIf(IsNull(g_rst_Princi!MAEETE_PORTEA), 0, g_rst_Princi!MAEETE_PORTEA), "0.00") & "  "
   End If
   If Not IsNull(g_rst_Princi!MAEETE_FECVCT) Then
      pnl_FecVct.Caption = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAEETE_FECVCT)), "dd/mm/yyyy") & "  "
   End If
   If Not IsNull(g_rst_Princi!MAEETE_FECAPR) Then
      pnl_FecApr.Caption = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAEETE_FECAPR)), "dd/mm/yyyy") & "  "
   End If
   If Not IsNull(g_rst_Princi!MAEETE_UBIGEO) Then
      Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!MAEETE_UBIGEO, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!MAEETE_UBIGEO, 2))
      Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!MAEETE_UBIGEO, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!MAEETE_UBIGEO, 2), Mid(g_rst_Princi!MAEETE_UBIGEO, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!MAEETE_UBIGEO, 2)))
   End If
   
   moddat_g_str_TipDoc = moddat_gf_Consulta_ParDes("118", g_rst_Princi!MAEETE_TIPDOC)
   moddat_g_str_NomCli = g_rst_Princi!MAEPRV_RAZSOC
   moddat_g_str_Descri = moddat_gf_Consulta_ParDes("526", g_rst_Princi!MAEETE_TIPEMP)
    
   l_int_TipDoc = moddat_g_int_TipDoc
   l_str_NumDoc = txt_NumDoc.Text
   l_dbl_LinAsi = Format(CStr(pnl_LinAsig_Ind.Caption), "###,###,###,##0.00")
   l_dbl_PorRet = CStr(Trim(pnl_PorRet.Caption))
   l_str_FecVct = Format(CStr(Trim(pnl_FecVct.Caption)), "yyyymmdd")
End Sub

Private Sub txt_ApeCas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_PriNom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeCas)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_CodSbs_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      If cmb_TipDoc.ListIndex <> -1 Then
         If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 1 Then
            Call gs_SetFocus(cmb_TipEmp)
         Else
            Call gs_SetFocus(txt_ApePat)
         End If
      End If
   End If
End Sub

Private Sub txt_DirRep_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
   End If
End Sub

Private Sub txt_NdoRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirRep)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-" & Chr(22))
   End If
End Sub

Private Sub txt_NomCor_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_ApePat)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmb_TipDoc.ListIndex <> -1 And Len(Trim(txt_NumDoc.Text)) > 0 Then
            Call fs_BuscarProv
            If cmb_TipDoc.ListIndex <> -1 Then
               If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) <> 1 Then
                  Call gs_SetFocus(txt_NomCor)
               Else
                  Call gs_SetFocus(txt_CodSbs)
               End If
            End If
        End If
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
    End If
End Sub

Private Sub fs_BuscarProv()
    pnl_RazSoc.Caption = ""
      
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " SELECT MAEPRV_TIPDOC, MAEPRV_NUMDOC, MAEPRV_RAZSOC, MAEPRV_CTADET, A.MAEPRV_TIPCNT,  "
    g_str_Parame = g_str_Parame & "        TRIM(B.PARDES_DESCRI) TIPCONTRIB, A.MAEPRV_CONDIC, TRIM(C.PARDES_DESCRI) CONDICION,  "
    g_str_Parame = g_str_Parame & "        MAEPRV_TIPPER, D.PARDES_DESCRI AS TIPOPERSONAL, MAEPRV_CODSIC  "
    g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A  "
    g_str_Parame = g_str_Parame & "         INNER JOIN MNT_PARDES B ON A.MAEPRV_TIPCNT = B.PARDES_CODITE AND B.PARDES_CODGRP = 119  "
    g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON A.MAEPRV_CONDIC = C.PARDES_CODITE AND C.PARDES_CODGRP = 120  "
    g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON A.MAEPRV_TIPPER = D.PARDES_CODITE AND D.PARDES_CODGRP = 127  "
    g_str_Parame = g_str_Parame & "  WHERE MAEPRV_SITUAC = 1  "
    
    If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & "  "
    End If
    If Len(Trim(txt_NumDoc.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEPRV_NUMDOC = '" & Trim(txt_NumDoc.Text) & "' "
    End If
    g_str_Parame = g_str_Parame & "  ORDER BY MAEPRV_TIPDOC, MAEPRV_RAZSOC ASC  "

    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
    
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       Screen.MousePointer = 0
       Exit Sub
    End If
    
    g_rst_Princi.MoveFirst
    If Not g_rst_Princi.EOF Then
       pnl_RazSoc.Caption = Trim(g_rst_Princi!MAEPRV_RAZSOC & "")
    End If
    
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
    Screen.MousePointer = 0
   
End Sub
Private Sub fs_Activa(ByVal p_Activa As Integer)
   
   cmb_TipEmp.Enabled = Not p_Activa
   pnl_LinAsig_Ind.Enabled = Not p_Activa
'   ipp_PorRet.Enabled = Not p_Activa
'   ipp_FecVct.Enabled = Not p_Activa
'   ipp_FecApr.Enabled = Not p_Activa
   cmd_Grabar.Enabled = Not p_Activa
   'cmd_Cancel.Enabled = Not p_Activa
   
   If moddat_g_int_FlgGrb = 1 Then
      cmb_TipDoc.Enabled = Not p_Activa
      txt_NumDoc.Enabled = Not p_Activa
      cmd_DatRCC.Enabled = p_Activa
      cmd_VtaPat.Enabled = p_Activa
      cmd_DatAcc.Enabled = p_Activa
      cmd_EteHis.Enabled = p_Activa
   Else
      cmb_TipDoc.Enabled = p_Activa
      txt_NumDoc.Enabled = p_Activa
   End If
End Sub
Private Sub fs_Limpia()
   
   cmb_TipDoc.ListIndex = -1
   cmb_TipEmp.ListIndex = -1
   cmb_CodCiu.ListIndex = -1
   cmb_TdoRep.ListIndex = -1
   
   txt_NumDoc.Text = Empty
   txt_NomCor.Text = Empty
   txt_CodSbs.Text = Empty
   txt_ApePat.Text = Empty
   txt_ApeMat.Text = Empty
   txt_PriNom.Text = Empty
   txt_SegNom.Text = Empty
   txt_ApeCas.Text = Empty
   txt_RepLeg.Text = Empty
   txt_NdoRep.Text = Empty
   txt_DirRep.Text = Empty
   txt_TelRep.Text = Empty
   
   pnl_RazSoc.Caption = Empty
   pnl_LinAsig_Ind.Caption = "0.00  "
   
'   ipp_AdmFlj.Text = 0#
'   ipp_ImpHip.Text = 0#
'   ipp_ImpLiq.Text = 0#
   
   pnl_PorRet.Caption = "0.00  "
   pnl_PorTEA.Caption = "0.00  "
   pnl_FecVct.Caption = ""
   pnl_FecApr.Caption = ""
   
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   
   l_int_TipDoc = 0
   l_str_NumDoc = ""
   l_dbl_LinAsi = 0
   l_dbl_PorRet = 0
   l_str_FecVct = ""
End Sub

Private Sub txt_PriNom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_SegNom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_RepLeg_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      Call gs_SetFocus(cmb_TdoRep)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
   End If
End Sub

Private Sub txt_SegNom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipEmp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_TelRep_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      Call gs_SetFocus(cmb_Situac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & " /-\" & Chr(22))
   End If
End Sub
