VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_SolCre_55 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10290
   ClientLeft      =   2445
   ClientTop       =   675
   ClientWidth     =   11520
   Icon            =   "OpeTra_frm_158.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10305
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   11535
      _Version        =   65536
      _ExtentX        =   20346
      _ExtentY        =   18177
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
         Height          =   2085
         Left            =   30
         TabIndex        =   38
         Top             =   5700
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   3678
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
         Begin VB.ComboBox cmb_FlgEst 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1710
            Width           =   885
         End
         Begin VB.ComboBox cmb_TipInm 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   720
            Width           =   1640
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   3720
            MaxLength       =   15
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   720
            Width           =   1665
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   8070
            TabIndex        =   14
            Text            =   "cmb_DptDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   2070
            TabIndex        =   15
            Text            =   "cmb_PrvDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   8070
            TabIndex        =   16
            Text            =   "cmb_DstDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   2070
            MaxLength       =   250
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.TextBox txt_Estaci 
            Height          =   315
            Left            =   9930
            MaxLength       =   120
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   1710
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Estac.:"
            Height          =   315
            Left            =   9030
            TabIndex        =   82
            Top             =   1740
            Width           =   945
         End
         Begin VB.Label Label49 
            Caption         =   "Tipo de Inmueble:"
            Height          =   315
            Left            =   90
            TabIndex        =   49
            Top             =   90
            Width           =   1365
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo de V�a:"
            Height          =   315
            Left            =   90
            TabIndex        =   48
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre V�a:"
            Height          =   285
            Left            =   6090
            TabIndex        =   47
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   90
            TabIndex        =   46
            Top             =   750
            Width           =   2055
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   6090
            TabIndex        =   45
            Top             =   750
            Width           =   1905
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   90
            TabIndex        =   44
            Top             =   1080
            Width           =   1485
         End
         Begin VB.Label Label24 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   6090
            TabIndex        =   43
            Top             =   1080
            Width           =   1665
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   90
            TabIndex        =   42
            Top             =   1410
            Width           =   1455
         End
         Begin VB.Label Label26 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   6090
            TabIndex        =   41
            Top             =   1410
            Width           =   1905
         End
         Begin VB.Label Label28 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   90
            TabIndex        =   40
            Top             =   1740
            Width           =   1485
         End
         Begin VB.Label Label47 
            Caption         =   "Estacionamiento:"
            Height          =   315
            Left            =   6090
            TabIndex        =   39
            Top             =   1740
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   735
         Left            =   30
         TabIndex        =   50
         Top             =   4140
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   1296
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Pro 
            Height          =   645
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   1138
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2415
         Left            =   30
         TabIndex        =   51
         Top             =   7830
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   4260
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
         Begin VB.TextBox txt_Telefo_Pro 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   250
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir_Pro 
            Height          =   315
            Left            =   8070
            TabIndex        =   31
            Text            =   "cmb_DstDir"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir_Pro 
            Height          =   315
            Left            =   2070
            TabIndex        =   30
            Text            =   "cmb_PrvDir"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir_Pro 
            Height          =   315
            Left            =   8070
            TabIndex        =   29
            Text            =   "cmb_DptDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon_Pro 
            Height          =   315
            Left            =   8070
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_IntDpt_Pro 
            Height          =   315
            Left            =   3720
            MaxLength       =   15
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   1050
            Width           =   1665
         End
         Begin VB.TextBox txt_NumVia_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   1050
            Width           =   1640
         End
         Begin VB.TextBox txt_NomVia_Pro 
            Height          =   315
            Left            =   8070
            MaxLength       =   120
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipVia_Pro 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NumDoc_Pro 
            Height          =   315
            Left            =   8070
            MaxLength       =   12
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipDoc_Pro 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_RazSoc_Pro 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   60
            Width           =   9345
         End
         Begin VB.Label Label29 
            Caption         =   "Tel�fono:"
            Height          =   285
            Left            =   6090
            TabIndex        =   64
            Top             =   2070
            Width           =   1725
         End
         Begin VB.Label Label27 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   90
            TabIndex        =   63
            Top             =   2070
            Width           =   1905
         End
         Begin VB.Label Label16 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   6090
            TabIndex        =   62
            Top             =   1740
            Width           =   1725
         End
         Begin VB.Label Label15 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   90
            TabIndex        =   61
            Top             =   1740
            Width           =   1905
         End
         Begin VB.Label Label14 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   6090
            TabIndex        =   60
            Top             =   1410
            Width           =   1725
         End
         Begin VB.Label Label13 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   90
            TabIndex        =   59
            Top             =   1410
            Width           =   1905
         End
         Begin VB.Label Label12 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   6090
            TabIndex        =   58
            Top             =   1080
            Width           =   1725
         End
         Begin VB.Label Label11 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   90
            TabIndex        =   57
            Top             =   1080
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Nombre V�a:"
            Height          =   285
            Left            =   6090
            TabIndex        =   56
            Top             =   750
            Width           =   1725
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo de V�a:"
            Height          =   315
            Left            =   90
            TabIndex        =   55
            Top             =   750
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "Nro. Doc. Identidad:"
            Height          =   285
            Left            =   6090
            TabIndex        =   54
            Top             =   420
            Width           =   1725
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   53
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label Label6 
            Caption         =   "Raz�n Social / Nombre:"
            Height          =   285
            Left            =   90
            TabIndex        =   52
            Top             =   90
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   1425
         Left            =   30
         TabIndex        =   65
         Top             =   2670
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   2514
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
         Begin VB.ComboBox cmb_Modali 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   9345
         End
         Begin VB.ComboBox cmb_Pryvin 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   9345
         End
         Begin VB.ComboBox cmb_Bancos 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   9345
         End
         Begin VB.ComboBox cmb_PryNVi 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1050
            Width           =   9345
         End
         Begin VB.Label Label48 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   69
            Top             =   90
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Proyecto Vinculado:"
            Height          =   315
            Left            =   90
            TabIndex        =   68
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label Label10 
            Caption         =   "Entidad Financiera:"
            Height          =   315
            Left            =   90
            TabIndex        =   67
            Top             =   750
            Width           =   1545
         End
         Begin VB.Label Label46 
            Caption         =   "Proyecto No Vinculado:"
            Height          =   315
            Left            =   90
            TabIndex        =   66
            Top             =   1080
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   70
         Top             =   2190
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
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
         Begin VB.ComboBox cmb_InmIde 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   1500
         End
         Begin VB.Label Label17 
            Caption         =   "Registra Inmueble:"
            Height          =   315
            Left            =   90
            TabIndex        =   71
            Top             =   90
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   72
         Top             =   690
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
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
         Begin VB.CommandButton cmd_SimCre 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_158.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Simulaci�n de Cr�ditos Hipotecarios"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_158.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10800
            Picture         =   "OpeTra_frm_158.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Salir de la Opci�n"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   615
         Left            =   30
         TabIndex        =   73
         Top             =   30
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   1085
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   255
            Left            =   690
            TabIndex        =   80
            Top             =   30
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Solicitud de Cr�dito Hipotecario"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   255
            Left            =   690
            TabIndex        =   81
            Top             =   300
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Datos del Inmueble"
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
            Picture         =   "OpeTra_frm_158.frx":0A62
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   74
         Top             =   1380
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   2070
            TabIndex        =   75
            Top             =   60
            Width           =   9315
            _Version        =   65536
            _ExtentX        =   16431
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   2070
            TabIndex        =   76
            Top             =   390
            Width           =   9315
            _Version        =   65536
            _ExtentX        =   16431
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   78
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   77
            Top             =   390
            Width           =   1755
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   735
         Left            =   30
         TabIndex        =   79
         Top             =   4920
         Width           =   11445
         _Version        =   65536
         _ExtentX        =   20188
         _ExtentY        =   1296
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Con 
            Height          =   645
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   1138
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_SolCre_55"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Bancos()      As moddat_tpo_Genera
Dim l_arr_PryVin()      As moddat_tpo_Genera
Dim l_arr_PryNVi()      As moddat_tpo_Genera
Dim l_arr_Modali()      As moddat_tpo_Genera
Dim l_str_DptDir_Pro    As String
Dim l_str_PrvDir_Pro    As String
Dim l_str_DstDir_Pro    As String
Dim l_str_DptDir        As String
Dim l_str_PrvDir        As String
Dim l_str_DstDir        As String
Dim l_int_FlgCmb        As Integer
Dim l_int_FlgAfeBV      As Integer
Dim l_int_FlgTipAfe     As Integer
Dim l_dbl_ValAfeBV      As Double

Private Sub cmb_Bancos_Click()
   If cmb_Bancos.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_PryNVi(cmb_PryNVi, l_arr_PryNVi, l_arr_Bancos(cmb_Bancos.ListIndex + 1).Genera_Codigo)
      Screen.MousePointer = 0
      
      Call gs_LimpiaGrid(grd_Listad_Pro)
      Call gs_LimpiaGrid(grd_Listad_Con)
      Call gs_SetFocus(cmb_PryNVi)
   End If
End Sub

Private Sub cmb_Bancos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Bancos_Click
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

Private Sub cmb_DptDir_Pro_Change()
   l_str_DptDir_Pro = cmb_DptDir_Pro.Text
End Sub

Private Sub cmb_DptDir_Pro_Click()
   If cmb_DptDir_Pro.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvDir_Pro.Clear
         cmb_DstDir_Pro.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir_Pro, Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir_Pro)
      End If
   End If
End Sub

Private Sub cmb_DptDir_Pro_GotFocus()
   Call SendMessage(cmb_DptDir_Pro.hwnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
   l_str_DptDir_Pro = cmb_DptDir_Pro.Text
End Sub

Private Sub cmb_DptDir_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir_Pro, l_str_DptDir_Pro)
      l_int_FlgCmb = True
      
      cmb_PrvDir_Pro.Clear
      cmb_DstDir_Pro.Clear
      If cmb_DptDir_Pro.ListIndex > -1 Then
         l_str_DptDir_Pro = ""
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir_Pro, Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir_Pro)
   End If
End Sub

Private Sub cmb_DptDir_Pro_LostFocus()
   Call SendMessage(cmb_DptDir_Pro.hwnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_DstDir_Change()
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_Click()
   If cmb_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Refere)
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
      
      Call gs_SetFocus(txt_Refere)
   End If
End Sub

Private Sub cmb_DstDir_LostFocus()
   Call SendMessage(cmb_DstDir.hwnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_FlgEst_Click()
   If cmb_FlgEst.ListIndex <> -1 Then
      If cmb_FlgEst.ItemData(cmb_FlgEst.ListIndex) = 1 Then
         txt_Estaci.Enabled = True
      Else
         txt_Estaci.Text = ""
         txt_Estaci.Enabled = False
      End If
   End If
End Sub

Private Sub cmb_FlgEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_FlgEst.ListIndex <> -1 Then
         If cmb_FlgEst.ItemData(cmb_FlgEst.ListIndex) = 1 Then
            Call gs_SetFocus(txt_Estaci)
         Else
            If cmb_Modali.ListIndex > -1 Then
               If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Or CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 4 Then
                  Call gs_SetFocus(txt_RazSoc_Pro)
               Else
                  Call gs_SetFocus(cmd_Grabar)
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub cmb_InmIde_Click()
   If cmb_InmIde.ListIndex > -1 Then
      If cmb_InmIde.ItemData(cmb_InmIde.ListIndex) = 1 Then
         cmb_Modali.Enabled = True
         cmb_Pryvin.Enabled = False
         cmb_Bancos.Enabled = False
         cmb_PryNVi.Enabled = False
         grd_Listad_Pro.Enabled = False
         grd_Listad_Con.Enabled = False
         cmb_TipInm.Enabled = True
         cmb_TipVia.Enabled = True
         txt_NomVia.Enabled = True
         txt_NumVia.Enabled = True
         txt_IntDpt.Enabled = True
         cmb_TipZon.Enabled = True
         txt_NomZon.Enabled = True
         cmb_DptDir.Enabled = True
         cmb_PrvDir.Enabled = True
         cmb_DstDir.Enabled = True
         txt_Refere.Enabled = True
         cmb_FlgEst.Enabled = True
         Call cmb_FlgEst_Click
         
         txt_RazSoc_Pro.Enabled = True
         cmb_TipDoc_Pro.Enabled = True
         cmb_TipVia_Pro.Enabled = True
         txt_NumDoc_Pro.Enabled = True
         txt_NomVia_Pro.Enabled = True
         txt_NumVia_Pro.Enabled = True
         txt_IntDpt_Pro.Enabled = True
         cmb_TipZon_Pro.Enabled = True
         txt_NomZon_Pro.Enabled = True
         cmb_DptDir_Pro.Enabled = True
         cmb_PrvDir_Pro.Enabled = True
         cmb_DstDir_Pro.Enabled = True
         txt_Refere_Pro.Enabled = True
         txt_Telefo_Pro.Enabled = True
         
         Call gs_SetFocus(cmb_Modali)
      Else
         cmb_Modali.ListIndex = -1
         cmb_Pryvin.ListIndex = -1
         cmb_Bancos.ListIndex = -1
         cmb_PryNVi.Clear
         Call gs_LimpiaGrid(grd_Listad_Pro)
         Call gs_LimpiaGrid(grd_Listad_Con)
         cmb_TipInm.ListIndex = -1
         cmb_TipVia.ListIndex = -1
         txt_NomVia.Text = ""
         txt_NumVia.Text = ""
         txt_IntDpt.Text = ""
         cmb_TipZon.ListIndex = -1
         txt_NomZon.Text = ""
         cmb_DptDir.ListIndex = -1
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         txt_Refere.Text = ""
         cmb_FlgEst.ListIndex = -1
         txt_Estaci.Text = ""
         
         txt_RazSoc_Pro.Text = ""
         cmb_TipDoc_Pro.ListIndex = -1
         txt_NumDoc_Pro.Text = ""
         cmb_TipVia_Pro.ListIndex = -1
         txt_NomVia_Pro.Text = ""
         txt_NumVia_Pro.Text = ""
         txt_IntDpt_Pro.Text = ""
         cmb_TipZon_Pro.ListIndex = -1
         txt_NomZon_Pro.Text = ""
         cmb_DptDir_Pro.ListIndex = -1
         cmb_PrvDir_Pro.Clear
         cmb_DstDir_Pro.Clear
         txt_Refere_Pro.Text = ""
         txt_Telefo_Pro.Text = ""
         
         cmb_Modali.Enabled = False
         cmb_Pryvin.Enabled = False
         cmb_Bancos.Enabled = False
         cmb_PryNVi.Enabled = False
         grd_Listad_Pro.Enabled = False
         grd_Listad_Con.Enabled = False
         cmb_TipInm.Enabled = False
         cmb_TipVia.Enabled = False
         txt_NomVia.Enabled = False
         txt_NumVia.Enabled = False
         txt_IntDpt.Enabled = False
         cmb_TipZon.Enabled = False
         txt_NomZon.Enabled = False
         cmb_DptDir.Enabled = False
         cmb_PrvDir.Enabled = False
         cmb_DstDir.Enabled = False
         txt_Refere.Enabled = False
         
         cmb_FlgEst.Enabled = False
         txt_Estaci.Enabled = False
         
         txt_RazSoc_Pro.Enabled = False
         cmb_TipDoc_Pro.Enabled = False
         txt_NumDoc_Pro.Enabled = False
         cmb_TipVia_Pro.Enabled = False
         txt_NomVia_Pro.Enabled = False
         txt_NumVia_Pro.Enabled = False
         txt_IntDpt_Pro.Enabled = False
         cmb_TipZon_Pro.Enabled = False
         txt_NomZon_Pro.Enabled = False
         cmb_DptDir_Pro.Enabled = False
         cmb_PrvDir_Pro.Enabled = False
         cmb_DstDir_Pro.Enabled = False
         txt_Refere_Pro.Enabled = False
         txt_Telefo_Pro.Enabled = False
      End If
   End If
End Sub

Private Sub cmb_Modali_Click()
   If cmb_Modali.ListIndex > -1 Then
      If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Or CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 4 Then
         cmb_Pryvin.Enabled = False
         cmb_Bancos.Enabled = False
         cmb_PryNVi.Enabled = False
         grd_Listad_Pro.Enabled = False
         grd_Listad_Con.Enabled = False
         cmb_Pryvin.ListIndex = -1
         cmb_Bancos.ListIndex = -1
         cmb_PryNVi.Clear
         Call gs_LimpiaGrid(grd_Listad_Pro)
         Call gs_LimpiaGrid(grd_Listad_Con)
         txt_RazSoc_Pro.Enabled = True
         cmb_TipDoc_Pro.Enabled = True
         txt_NumDoc_Pro.Enabled = True
         cmb_TipVia_Pro.Enabled = True
         txt_NomVia_Pro.Enabled = True
         txt_NumVia_Pro.Enabled = True
         txt_IntDpt_Pro.Enabled = True
         cmb_TipZon_Pro.Enabled = True
         txt_NomZon_Pro.Enabled = True
         cmb_DptDir_Pro.Enabled = True
         cmb_PrvDir_Pro.Enabled = True
         cmb_DstDir_Pro.Enabled = True
         txt_Refere_Pro.Enabled = True
         txt_Telefo_Pro.Enabled = True
         Call gs_SetFocus(cmb_TipInm)
         
      ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Then
         cmb_Pryvin.Enabled = False
         cmb_Pryvin.ListIndex = -1
         cmb_Bancos.Enabled = True
         cmb_PryNVi.Enabled = True
         Call gs_LimpiaGrid(grd_Listad_Pro)
         Call gs_LimpiaGrid(grd_Listad_Con)
         grd_Listad_Pro.Enabled = True
         grd_Listad_Con.Enabled = True
         cmb_PryNVi.Clear
         txt_RazSoc_Pro.Text = ""
         cmb_TipDoc_Pro.ListIndex = -1
         txt_NumDoc_Pro.Text = ""
         cmb_TipVia_Pro.ListIndex = -1
         txt_NomVia_Pro.Text = ""
         txt_NumVia_Pro.Text = ""
         txt_IntDpt_Pro.Text = ""
         cmb_TipZon_Pro.ListIndex = -1
         txt_NomZon_Pro.Text = ""
         cmb_DptDir_Pro.ListIndex = -1
         cmb_PrvDir_Pro.Clear
         cmb_DstDir_Pro.Clear
         txt_Refere_Pro.Text = ""
         txt_Telefo_Pro.Text = ""
         txt_RazSoc_Pro.Enabled = False
         cmb_TipDoc_Pro.Enabled = False
         txt_NumDoc_Pro.Enabled = False
         cmb_TipVia_Pro.Enabled = False
         txt_NomVia_Pro.Enabled = False
         txt_NumVia_Pro.Enabled = False
         txt_IntDpt_Pro.Enabled = False
         cmb_TipZon_Pro.Enabled = False
         txt_NomZon_Pro.Enabled = False
         cmb_DptDir_Pro.Enabled = False
         cmb_PrvDir_Pro.Enabled = False
         cmb_DstDir_Pro.Enabled = False
         txt_Refere_Pro.Enabled = False
         txt_Telefo_Pro.Enabled = False
         Call gs_SetFocus(cmb_Bancos)
         
      ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Then
         cmb_Pryvin.Enabled = True
         cmb_Bancos.Enabled = False
         cmb_PryNVi.Enabled = False
         Call gs_LimpiaGrid(grd_Listad_Pro)
         Call gs_LimpiaGrid(grd_Listad_Con)
         grd_Listad_Pro.Enabled = True
         grd_Listad_Con.Enabled = True
         cmb_Bancos.ListIndex = -1
         cmb_PryNVi.Clear
         txt_RazSoc_Pro.Text = ""
         cmb_TipDoc_Pro.ListIndex = -1
         txt_NumDoc_Pro.Text = ""
         cmb_TipVia_Pro.ListIndex = -1
         txt_NomVia_Pro.Text = ""
         txt_NumVia_Pro.Text = ""
         txt_IntDpt_Pro.Text = ""
         cmb_TipZon_Pro.ListIndex = -1
         txt_NomZon_Pro.Text = ""
         cmb_DptDir_Pro.ListIndex = -1
         cmb_PrvDir_Pro.Clear
         cmb_DstDir_Pro.Clear
         txt_Refere_Pro.Text = ""
         txt_Telefo_Pro.Text = ""
         txt_RazSoc_Pro.Enabled = False
         cmb_TipDoc_Pro.Enabled = False
         txt_NumDoc_Pro.Enabled = False
         cmb_TipVia_Pro.Enabled = False
         txt_NomVia_Pro.Enabled = False
         txt_NumVia_Pro.Enabled = False
         txt_IntDpt_Pro.Enabled = False
         cmb_TipZon_Pro.Enabled = False
         txt_NomZon_Pro.Enabled = False
         cmb_DptDir_Pro.Enabled = False
         cmb_PrvDir_Pro.Enabled = False
         cmb_DstDir_Pro.Enabled = False
         txt_Refere_Pro.Enabled = False
         txt_Telefo_Pro.Enabled = False
         Call gs_SetFocus(cmb_Pryvin)
         
      End If
   End If
End Sub

Private Sub cmb_Modali_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Modali_Click
   End If
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

Private Sub cmb_PrvDir_Pro_Change()
   l_str_PrvDir_Pro = cmb_PrvDir_Pro.Text
End Sub

Private Sub cmb_PrvDir_Pro_Click()
   If cmb_PrvDir_Pro.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstDir_Pro.Clear
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir_Pro, Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00"), Format(cmb_PrvDir_Pro.ItemData(cmb_PrvDir_Pro.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir_Pro)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_Pro_GotFocus()
   Call SendMessage(cmb_PrvDir_Pro.hwnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
   l_str_PrvDir_Pro = cmb_PrvDir_Pro.Text
End Sub

Private Sub cmb_PrvDir_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir_Pro, l_str_PrvDir_Pro)
      l_int_FlgCmb = True
      
      cmb_DstDir_Pro.Clear
      If cmb_PrvDir_Pro.ListIndex > -1 Then
         l_str_DstDir_Pro = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir_Pro, Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00"), Format(cmb_PrvDir_Pro.ItemData(cmb_PrvDir_Pro.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir_Pro)
   End If
End Sub

Private Sub cmb_PrvDir_Pro_LostFocus()
   Call SendMessage(cmb_PrvDir_Pro.hwnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_DstDir_Pro_Change()
   l_str_DstDir_Pro = cmb_DstDir_Pro.Text
End Sub

Private Sub cmb_DstDir_Pro_Click()
   If cmb_DstDir_Pro.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Refere_Pro)
      End If
   End If
End Sub

Private Sub cmb_DstDir_Pro_GotFocus()
   Call SendMessage(cmb_DstDir_Pro.hwnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
   l_str_DstDir_Pro = cmb_DstDir_Pro.Text
End Sub

Private Sub cmb_DstDir_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir_Pro, l_str_DstDir_Pro)
      l_int_FlgCmb = True
      
      If cmb_DstDir_Pro.ListIndex > -1 Then
         l_str_DstDir_Pro = ""
      End If
      
      Call gs_SetFocus(txt_Refere_Pro)
   End If
End Sub

Private Sub cmb_DstDir_Pro_LostFocus()
   Call SendMessage(cmb_DstDir_Pro.hwnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_PryNVi_Click()
   If cmb_PryNVi.ListIndex > -1 Then
      Call gs_LimpiaGrid(grd_Listad_Pro)
      Call gs_LimpiaGrid(grd_Listad_Con)
   
      grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
      grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
      grd_Listad_Pro.Col = 0
      grd_Listad_Pro.Text = "Doc. Ident. Promotor"
      
      grd_Listad_Pro.Col = 1
      grd_Listad_Pro.Text = CStr(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenTDo) & "-" & Trim(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenNDo)
      
      grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
      grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
      grd_Listad_Pro.Col = 0
      grd_Listad_Pro.Text = "Raz�n Social Promotor"
      
      grd_Listad_Pro.Col = 1
      grd_Listad_Pro.Text = moddat_gf_Consulta_RazSoc(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenTDo, l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenNDo)
      
      'Constructor
      grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
      grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
      grd_Listad_Con.Col = 0
      grd_Listad_Con.Text = "Doc. Ident. Constructor"
      
      grd_Listad_Con.Col = 1
      grd_Listad_Con.Text = CStr(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConTDo) & "-" & Trim(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConNDo)
      
      grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
      grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
      grd_Listad_Con.Col = 0
      grd_Listad_Con.Text = "Raz�n Social Constructor"
      
      grd_Listad_Con.Col = 1
      grd_Listad_Con.Text = moddat_gf_Consulta_RazSoc(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConTDo, l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConNDo)
      
      Call gs_BuscarCombo_Item(cmb_TipVia, l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_TipVia)
      txt_NomVia.Text = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_NomVia
      txt_NumVia.Text = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_NumVia
      txt_IntDpt.Text = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_IntDpt
      Call gs_BuscarCombo_Item(cmb_TipZon, l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_TipZon)
      txt_NomZon.Text = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_NomZon
      Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 2))
      Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstDir, Left(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 2), Mid(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_UbiGeo, 2)))
      txt_Refere.Text = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_Refere
      
      Call gs_UbiIniGrid(grd_Listad_Pro)
      Call gs_UbiIniGrid(grd_Listad_Con)
      
      Call gs_SetFocus(cmb_TipInm)
   End If
End Sub

Private Sub cmb_PryNVi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PryNVi_Click
   End If
End Sub

Private Sub cmb_Pryvin_Click()
   If cmb_Pryvin.ListIndex > -1 Then
      Call gs_LimpiaGrid(grd_Listad_Pro)
      Call gs_LimpiaGrid(grd_Listad_Con)
      
      grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
      grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
      grd_Listad_Pro.Col = 0
      grd_Listad_Pro.Text = "Doc. Ident. Promotor"
      
      grd_Listad_Pro.Col = 1
      grd_Listad_Pro.Text = CStr(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenTDo) & "-" & Trim(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenNDo)
      
      grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
      grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
      grd_Listad_Pro.Col = 0
      grd_Listad_Pro.Text = "Raz�n Social Promotor"
      
      grd_Listad_Pro.Col = 1
      grd_Listad_Pro.Text = moddat_gf_Consulta_RazSoc(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenTDo, l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenNDo)
      
      'Constructor
      grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
      grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
      grd_Listad_Con.Col = 0
      grd_Listad_Con.Text = "Doc. Ident. Constructor"
      
      grd_Listad_Con.Col = 1
      grd_Listad_Con.Text = CStr(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConTDo) & "-" & Trim(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConNDo)
      
      grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
      grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
      grd_Listad_Con.Col = 0
      grd_Listad_Con.Text = "Raz�n Social Constructor"
      
      grd_Listad_Con.Col = 1
      grd_Listad_Con.Text = moddat_gf_Consulta_RazSoc(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConTDo, l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConNDo)
      
      Call gs_BuscarCombo_Item(cmb_TipVia, l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_TipVia)
      txt_NomVia.Text = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_NomVia
      txt_NumVia.Text = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_NumVia
      txt_IntDpt.Text = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_IntDpt
      Call gs_BuscarCombo_Item(cmb_TipZon, l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_TipZon)
      txt_NomZon.Text = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_NomZon
      Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 2))
      Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstDir, Left(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 2), Mid(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_UbiGeo, 2)))
      txt_Refere.Text = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_Refere
      
      Call gs_UbiIniGrid(grd_Listad_Pro)
      Call gs_UbiIniGrid(grd_Listad_Con)
      
      Call gs_SetFocus(cmb_TipInm)
   End If
End Sub

Private Sub cmb_Pryvin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Pryvin_Click
   End If
End Sub

Private Sub cmb_TipDoc_Pro_Click()
   If cmb_TipDoc_Pro.ListIndex > -1 Then
      Select Case cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex)
         Case 1:  txt_NumDoc_Pro.MaxLength = 8
         Case 7:  txt_NumDoc_Pro.MaxLength = 11
         Case Else:  txt_NumDoc_Pro.MaxLength = 12
      End Select
   End If
   
   Call gs_SetFocus(txt_NumDoc_Pro)
End Sub

Private Sub cmb_TipDoc_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Pro_Click
   End If
End Sub

Private Sub cmb_TipInm_Click()
   Call gs_SetFocus(cmb_TipVia)
End Sub

Private Sub cmb_TipInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipInm_Click
   End If
End Sub

Private Sub cmb_TipVia_Click()
   Call gs_SetFocus(txt_NomVia)
End Sub

Private Sub cmb_TipVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipVia_Click
   End If
End Sub

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipZon_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_CodPry As String

   If cmb_InmIde.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cliente registra Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_InmIde)
      Exit Sub
   End If
   
   If cmb_InmIde.ItemData(cmb_InmIde.ListIndex) = 1 Then
      If cmb_Modali.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Modalidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Modali)
         Exit Sub
      End If
      
      If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Then
         If cmb_Bancos.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Entidad Financiera.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Bancos)
            Exit Sub
         End If
         If cmb_PryNVi.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Proyecto No Vinculado.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_PryNVi)
            Exit Sub
         End If
      ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Then
         If cmb_Pryvin.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Proyecto Vinculado.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Pryvin)
            Exit Sub
         End If
      End If
      If cmb_TipInm.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Inmueble.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipInm)
         Exit Sub
      End If
      If cmb_TipVia.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de V�a.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipVia)
         Exit Sub
      End If
      If Len(Trim(txt_NomVia.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de V�a.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomVia)
         Exit Sub
      End If
      If Len(Trim(txt_NumVia.Text)) = 0 Then
         MsgBox "Debe ingresar el N�mero.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumVia)
         Exit Sub
      End If
      If cmb_TipZon.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipZon)
         Exit Sub
      End If
      If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) <> 12 Then
         If Len(Trim(txt_NomZon.Text)) = 0 Then
            MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomZon)
            Exit Sub
         End If
      End If
      If cmb_DptDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Departamento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DptDir)
         Exit Sub
      End If
      If cmb_PrvDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvDir)
         Exit Sub
      End If
      If cmb_DstDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Distrito.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DstDir)
         Exit Sub
      End If
      
      If cmb_FlgEst.ListIndex = -1 Then
         MsgBox "Debe seleccionar si tiene estacionamiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_FlgEst)
         Exit Sub
      End If
      
      If cmb_FlgEst.ItemData(cmb_FlgEst.ListIndex) = 1 Then
         If Trim(txt_Estaci.Text & "") = "" Then
            MsgBox "Debe de ingresar el numero de estacionamiento.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Estaci)
            Exit Sub
         End If
      End If
      
      If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Or CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 4 Then
         If Len(Trim(txt_RazSoc_Pro.Text)) = 0 Then
            MsgBox "Debe ingresar el Nombre o Raz�n Social.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_RazSoc_Pro)
            Exit Sub
         End If
         If cmb_TipDoc_Pro.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TipDoc_Pro)
            Exit Sub
         End If
         If Len(Trim(txt_NumDoc_Pro.Text)) = 0 Then
            MsgBox "Debe ingresar el N�mero de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumDoc_Pro)
            Exit Sub
         End If
         If cmb_TipVia_Pro.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de V�a.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TipVia_Pro)
            Exit Sub
         End If
         If Len(Trim(txt_NomVia_Pro.Text)) = 0 Then
            MsgBox "Debe ingresar el Nombre de V�a.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomVia_Pro)
            Exit Sub
         End If
         If Len(Trim(txt_NumVia_Pro.Text)) = 0 Then
            MsgBox "Debe ingresar el N�mero.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumVia_Pro)
            Exit Sub
         End If
         If cmb_TipZon_Pro.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TipZon_Pro)
            Exit Sub
         End If
         If cmb_TipZon_Pro.ItemData(cmb_TipZon_Pro.ListIndex) <> 12 Then
            If Len(Trim(txt_NomZon_Pro.Text)) = 0 Then
               MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NomZon_Pro)
               Exit Sub
            End If
         End If
         If cmb_DptDir_Pro.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Departamento.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_DptDir_Pro)
            Exit Sub
         End If
         If cmb_PrvDir_Pro.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Provincia.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_PrvDir_Pro)
            Exit Sub
         End If
         If cmb_DstDir_Pro.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Distrito.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_DstDir_Pro)
            Exit Sub
         End If
      End If
   End If
   
   'Valida el BMS en la informacion del credito
   If cmb_Modali.ListIndex <> -1 Then
      If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Then
         r_str_CodPry = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_Codigo
      ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Then
         r_str_CodPry = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_Codigo
      Else
         r_str_CodPry = 0
      End If
      
      moddat_g_str_Codigo = r_str_CodPry
      
      If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 And r_str_CodPry <> 0 Then
         Call moddat_gs_Consulta_DatBMS(r_str_CodPry, l_int_FlgAfeBV, l_int_FlgTipAfe, l_dbl_ValAfeBV)
          
         If l_int_FlgAfeBV = 1 And modatecli_g_arr_DatCre(1).DatCre_MtoBMS_Sol = 0 Then
            modatecli_g_int_DatCreTit = 1
            MsgBox "Proyecto seleccionado esta afecto al BMS, favor verificar informaci�n del cr�dito.", vbExclamation, modgen_g_str_NomPlt
         ElseIf l_int_FlgAfeBV = 0 And modatecli_g_arr_DatCre(1).DatCre_MtoBMS_Sol > 0 Then
            modatecli_g_int_DatCreTit = 1
            MsgBox "Proyecto seleccionado no esta afecta al BMS, favor verificar informaci�n del cr�dito.", vbExclamation, modgen_g_str_NomPlt
         End If
      End If
   End If
   If MsgBox("�Est� seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Call modatecli_gs_Limpia_DatInm
   
   'Pasar informaci�n al Arreglo
   modatecli_g_arr_DatInm(1).DatInm_InmIde = cmb_InmIde.ItemData(cmb_InmIde.ListIndex)
   
   If cmb_InmIde.ItemData(cmb_InmIde.ListIndex) = 1 Then
      modatecli_g_arr_DatInm(1).DatInm_TipInm = cmb_TipInm.ItemData(cmb_TipInm.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_Modali = l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo
      modatecli_g_arr_DatInm(1).DatInm_TipVia = cmb_TipVia.ItemData(cmb_TipVia.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_NomVia = txt_NomVia.Text
      modatecli_g_arr_DatInm(1).DatInm_Numero = txt_NumVia.Text
      modatecli_g_arr_DatInm(1).DatInm_Interi = txt_IntDpt.Text
      modatecli_g_arr_DatInm(1).DatInm_TipZon = cmb_TipZon.ItemData(cmb_TipZon.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_NomZon = txt_NomZon.Text
      modatecli_g_arr_DatInm(1).DatInm_UbiGeo = Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00")
      modatecli_g_arr_DatInm(1).DatInm_Refere = txt_Refere.Text
      modatecli_g_arr_DatInm(1).DatInm_FlgEst = cmb_FlgEst.ItemData(cmb_FlgEst.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_Estaci = Trim(txt_Estaci.Text)
      
      If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Then
         modatecli_g_arr_DatInm(1).DatInm_PryMCs = 1
         modatecli_g_arr_DatInm(1).DatInm_CodPry = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_Codigo
         modatecli_g_arr_DatInm(1).DatInm_FlgPro = 2
         modatecli_g_arr_DatInm(1).DatInm_TipDoc_Pro = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenTDo
         modatecli_g_arr_DatInm(1).DatInm_NumDoc_Pro = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_VenNDo
         modatecli_g_arr_DatInm(1).DatInm_FlgCon = 1
         modatecli_g_arr_DatInm(1).DatInm_TipDoc_Con = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConTDo
         modatecli_g_arr_DatInm(1).DatInm_NumDoc_Con = l_arr_PryVin(cmb_Pryvin.ListIndex + 1).Genera_ConNDo
      
      ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Then
         modatecli_g_arr_DatInm(1).DatInm_PryMCs = 2
         modatecli_g_arr_DatInm(1).DatInm_BcoPry = l_arr_Bancos(cmb_Bancos.ListIndex + 1).Genera_Codigo
         modatecli_g_arr_DatInm(1).DatInm_CodPry = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_Codigo
         modatecli_g_arr_DatInm(1).DatInm_FlgPro = 2
         modatecli_g_arr_DatInm(1).DatInm_TipDoc_Pro = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenTDo
         modatecli_g_arr_DatInm(1).DatInm_NumDoc_Pro = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_VenNDo
         modatecli_g_arr_DatInm(1).DatInm_FlgCon = 1
         modatecli_g_arr_DatInm(1).DatInm_TipDoc_Con = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConTDo
         modatecli_g_arr_DatInm(1).DatInm_NumDoc_Con = l_arr_PryNVi(cmb_PryNVi.ListIndex + 1).Genera_ConNDo
         
      Else
         modatecli_g_arr_DatInm(1).DatInm_PryMCs = 2
         modatecli_g_arr_DatInm(1).DatInm_FlgPro = 1
         modatecli_g_arr_DatInm(1).DatInm_TipDoc_Pro = cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_NumDoc_Pro = txt_NumDoc_Pro.Text
         modatecli_g_arr_DatInm(1).DatInm_RazSoc_Pro = txt_RazSoc_Pro.Text
         modatecli_g_arr_DatInm(1).DatInm_TipVia_Pro = cmb_TipVia_Pro.ItemData(cmb_TipVia_Pro.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_NomVia_Pro = txt_NomVia_Pro.Text
         modatecli_g_arr_DatInm(1).DatInm_NumVia_Pro = txt_NumVia_Pro.Text
         modatecli_g_arr_DatInm(1).DatInm_IntDpt_Pro = txt_IntDpt_Pro.Text
         modatecli_g_arr_DatInm(1).DatInm_TipZon_Pro = cmb_TipZon_Pro.ItemData(cmb_TipZon_Pro.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_NomZon_Pro = txt_NomZon_Pro.Text
         modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Pro = Format(cmb_DptDir_Pro.ItemData(cmb_DptDir_Pro.ListIndex), "00") & Format(cmb_PrvDir_Pro.ItemData(cmb_PrvDir_Pro.ListIndex), "00") & Format(cmb_DstDir_Pro.ItemData(cmb_DstDir_Pro.ListIndex), "00")
         modatecli_g_arr_DatInm(1).DatInm_Refere_Pro = txt_Refere_Pro.Text
         modatecli_g_arr_DatInm(1).DatInm_Telefo_Pro = txt_Telefo_Pro.Text
         
      End If
      modatecli_g_arr_DatInm(1).DatInm_NomPry = ""
   End If
   
   modatecli_g_int_DatInmTit = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opci�n.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_11.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_Produc.Caption = moddat_gf_Consulta_Produc(moddat_g_str_CodPrd)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Limpia
   
   If modatecli_g_int_DatInmTit = 2 Then
      Call gs_BuscarCombo_Item(cmb_InmIde, modatecli_g_arr_DatInm(1).DatInm_InmIde)
      If cmb_InmIde.ItemData(cmb_InmIde.ListIndex) = 1 Then
         cmb_Modali.ListIndex = gf_Busca_Arregl(l_arr_Modali, modatecli_g_arr_DatInm(1).DatInm_Modali) - 1
         
         If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 2 Then
            cmb_Bancos.ListIndex = gf_Busca_Arregl(l_arr_Bancos, modatecli_g_arr_DatInm(1).DatInm_BcoPry) - 1
            cmb_PryNVi.ListIndex = gf_Busca_Arregl(l_arr_PryNVi, modatecli_g_arr_DatInm(1).DatInm_CodPry) - 1
            cmb_Bancos.Enabled = True
            cmb_PryNVi.Enabled = True
            grd_Listad_Pro.Enabled = True
            grd_Listad_Con.Enabled = True
         ElseIf CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 3 Then
            cmb_Pryvin.ListIndex = gf_Busca_Arregl(l_arr_PryVin, modatecli_g_arr_DatInm(1).DatInm_CodPry) - 1
            cmb_Pryvin.Enabled = True
            grd_Listad_Pro.Enabled = True
            grd_Listad_Con.Enabled = True
         End If
         
         Call gs_BuscarCombo_Item(cmb_TipInm, modatecli_g_arr_DatInm(1).DatInm_TipInm)
         Call gs_BuscarCombo_Item(cmb_TipVia, modatecli_g_arr_DatInm(1).DatInm_TipVia)
         txt_NomVia.Text = modatecli_g_arr_DatInm(1).DatInm_NomVia
         txt_NumVia.Text = modatecli_g_arr_DatInm(1).DatInm_Numero
         txt_IntDpt.Text = modatecli_g_arr_DatInm(1).DatInm_Interi
         Call gs_BuscarCombo_Item(cmb_TipZon, modatecli_g_arr_DatInm(1).DatInm_TipZon)
         txt_NomZon.Text = modatecli_g_arr_DatInm(1).DatInm_NomZon
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(modatecli_g_arr_DatInm(1).DatInm_UbiGeo, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(modatecli_g_arr_DatInm(1).DatInm_UbiGeo, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_UbiGeo, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(modatecli_g_arr_DatInm(1).DatInm_UbiGeo, 2), Mid(modatecli_g_arr_DatInm(1).DatInm_UbiGeo, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(modatecli_g_arr_DatInm(1).DatInm_UbiGeo, 2)))
         txt_Refere.Text = modatecli_g_arr_DatInm(1).DatInm_Refere
         Call gs_BuscarCombo_Item(cmb_FlgEst, modatecli_g_arr_DatInm(1).DatInm_FlgEst)
         txt_Estaci.Text = Trim(modatecli_g_arr_DatInm(1).DatInm_Estaci & "")
         
         If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Or CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 4 Then
            txt_RazSoc_Pro.Text = modatecli_g_arr_DatInm(1).DatInm_RazSoc_Pro
            Call gs_BuscarCombo_Item(cmb_TipDoc_Pro, modatecli_g_arr_DatInm(1).DatInm_TipDoc_Pro)
            txt_NumDoc_Pro.Text = modatecli_g_arr_DatInm(1).DatInm_NumDoc_Pro
            Call gs_BuscarCombo_Item(cmb_TipVia_Pro, modatecli_g_arr_DatInm(1).DatInm_TipVia_Pro)
            txt_NomVia_Pro.Text = modatecli_g_arr_DatInm(1).DatInm_NomVia_Pro
            txt_NumVia_Pro.Text = modatecli_g_arr_DatInm(1).DatInm_NumVia_Pro
            txt_IntDpt_Pro.Text = modatecli_g_arr_DatInm(1).DatInm_IntDpt_Pro
            Call gs_BuscarCombo_Item(cmb_TipZon_Pro, modatecli_g_arr_DatInm(1).DatInm_TipZon_Pro)
            txt_NomZon_Pro.Text = modatecli_g_arr_DatInm(1).DatInm_NomZon_Pro
            Call gs_BuscarCombo_Item(cmb_DptDir_Pro, CInt(Left(modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Pro, 2)))
            Call moddat_gs_Carga_Provin(cmb_PrvDir_Pro, Left(modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Pro, 2))
            Call gs_BuscarCombo_Item(cmb_PrvDir_Pro, CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Pro, 3, 2)))
            Call moddat_gs_Carga_Distri(cmb_DstDir_Pro, Left(modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Pro, 2), Mid(modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Pro, 3, 2))
            Call gs_BuscarCombo_Item(cmb_DstDir_Pro, CInt(Right(modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Pro, 2)))
            txt_Refere_Pro.Text = modatecli_g_arr_DatInm(1).DatInm_Refere_Pro
            txt_Telefo_Pro.Text = modatecli_g_arr_DatInm(1).DatInm_Telefo_Pro
         End If
      End If
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_InmIde, 1, "214")
   Call moddat_gs_Carga_ParSubPrd(cmb_Modali, l_arr_Modali(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003")
   Call moddat_gs_Carga_LisIte(cmb_Bancos, l_arr_Bancos, 1, "513")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipInm, 1, "217")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst, 1, "214")
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   Call moddat_gs_Carga_PryVin(cmb_Pryvin, l_arr_PryVin)
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Pro, 1, "236")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia_Pro, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon_Pro, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir_Pro)

   grd_Listad_Pro.ColWidth(0) = 3200
   grd_Listad_Pro.ColWidth(1) = 7700
   grd_Listad_Pro.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_Pro.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Listad_Con.ColWidth(0) = 3200
   grd_Listad_Con.ColWidth(1) = 7700
   grd_Listad_Con.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_Con.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Limpia()
   cmb_InmIde.ListIndex = -1
   cmb_Modali.ListIndex = -1
   cmb_Pryvin.ListIndex = -1
   cmb_Bancos.ListIndex = -1
   cmb_PryNVi.Clear
   Call gs_LimpiaGrid(grd_Listad_Pro)
   Call gs_LimpiaGrid(grd_Listad_Con)
   cmb_TipInm.ListIndex = -1
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_NumVia.Text = ""
   txt_IntDpt.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   cmb_FlgEst.ListIndex = -1
   txt_Estaci.Text = ""
   
   txt_RazSoc_Pro.Text = ""
   cmb_TipDoc_Pro.ListIndex = -1
   txt_NumDoc_Pro.Text = ""
   cmb_TipVia_Pro.ListIndex = -1
   txt_NomVia_Pro.Text = ""
   txt_NumVia_Pro.Text = ""
   txt_IntDpt_Pro.Text = ""
   cmb_TipZon_Pro.ListIndex = -1
   txt_NomZon_Pro.Text = ""
   cmb_DptDir_Pro.ListIndex = -1
   cmb_PrvDir_Pro.Clear
   cmb_DstDir_Pro.Clear
   txt_Refere_Pro.Text = ""
   txt_Telefo_Pro.Text = ""
   cmb_Modali.Enabled = False
   cmb_Pryvin.Enabled = False
   cmb_Bancos.Enabled = False
   cmb_PryNVi.Enabled = False
   grd_Listad_Pro.Enabled = False
   grd_Listad_Con.Enabled = False
   cmb_TipInm.Enabled = False
   cmb_TipVia.Enabled = False
   txt_NomVia.Enabled = False
   txt_NumVia.Enabled = False
   txt_IntDpt.Enabled = False
   cmb_TipZon.Enabled = False
   txt_NomZon.Enabled = False
   cmb_DptDir.Enabled = False
   cmb_PrvDir.Enabled = False
   cmb_DstDir.Enabled = False
   txt_Refere.Enabled = False
   cmb_FlgEst.Enabled = False
   txt_Estaci.Enabled = False
   txt_RazSoc_Pro.Enabled = False
   cmb_TipDoc_Pro.Enabled = False
   txt_NumDoc_Pro.Enabled = False
   cmb_TipVia_Pro.Enabled = False
   txt_NomVia_Pro.Enabled = False
   txt_NumVia_Pro.Enabled = False
   txt_IntDpt_Pro.Enabled = False
   cmb_TipZon_Pro.Enabled = False
   txt_NomZon_Pro.Enabled = False
   cmb_DptDir_Pro.Enabled = False
   cmb_PrvDir_Pro.Enabled = False
   cmb_DstDir_Pro.Enabled = False
   txt_Refere_Pro.Enabled = False
   txt_Telefo_Pro.Enabled = False
End Sub

Private Sub grd_Listad_Con_SelChange()
   If grd_Listad_Con.Rows > 2 Then
      grd_Listad_Con.RowSel = grd_Listad_Con.Row
   End If
End Sub

Private Sub grd_Listad_Pro_SelChange()
   If grd_Listad_Pro.Rows > 2 Then
      grd_Listad_Pro.RowSel = grd_Listad_Pro.Row
   End If
End Sub

Private Sub txt_Estaci_GotFocus()
   Call gs_SelecTodo(txt_Estaci)
End Sub

Private Sub txt_Estaci_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_Modali.ListIndex > -1 Then
         If CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 1 Or CInt(l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo) = 4 Then
            Call gs_SetFocus(txt_RazSoc_Pro)
         Else
            Call gs_SetFocus(cmd_Grabar)
         End If
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?��")
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NomZon_GotFocus()
   Call gs_SelecTodo(txt_NomZon)
End Sub

Private Sub txt_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NumDoc_Pro_GotFocus()
   Call gs_SelecTodo(txt_NumDoc_Pro)
End Sub

Private Sub txt_NumDoc_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia_Pro)
   Else
      If cmb_TipDoc_Pro.ListIndex > -1 Then
         Select Case cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex)
            Case 1:     KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 7:     KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case Else:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub cmb_TipVia_Pro_Click()
   Call gs_SetFocus(txt_NomVia_Pro)
End Sub

Private Sub cmb_TipVia_Pro_KeyPress(KeyAscii As Integer)
   Call cmb_TipVia_Pro_Click
End Sub

Private Sub cmb_TipZon_Pro_Click()
   Call gs_SetFocus(txt_NomZon_Pro)
End Sub

Private Sub cmb_TipZon_Pro_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Pro_Click
End Sub

Private Sub txt_NomVia_Pro_GotFocus()
   Call gs_SelecTodo(txt_NomVia_Pro)
End Sub

Private Sub txt_NomVia_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NomZon_Pro_GotFocus()
   Call gs_SelecTodo(txt_NomZon_Pro)
End Sub

Private Sub txt_NomZon_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NumVia_Pro_GotFocus()
   Call gs_SelecTodo(txt_NumVia_Pro)
End Sub

Private Sub txt_NumVia_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_IntDpt_Pro_GotFocus()
   Call gs_SelecTodo(txt_IntDpt_Pro)
End Sub

Private Sub txt_IntDpt_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NumVia_GotFocus()
   Call gs_SelecTodo(txt_NumVia)
End Sub

Private Sub txt_NumVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_IntDpt_GotFocus()
   Call gs_SelecTodo(txt_IntDpt)
End Sub

Private Sub txt_IntDpt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_RazSoc_Pro_GotFocus()
   Call gs_SelecTodo(txt_RazSoc_Pro)
End Sub

Private Sub txt_RazSoc_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ',;:.)(@#$%&/?�_")
   End If
End Sub

Private Sub txt_Refere_Pro_GotFocus()
   Call gs_SelecTodo(txt_Refere_Pro)
End Sub

Private Sub txt_Refere_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telefo_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?��")
   End If
End Sub

Private Sub txt_Telefo_Pro_GotFocus()
   Call gs_SelecTodo(txt_Telefo_Pro)
End Sub

Private Sub txt_Telefo_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_FlgEst)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?��")
   End If
End Sub
