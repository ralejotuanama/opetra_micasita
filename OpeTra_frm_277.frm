VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Tra_EvaTas_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6840
   ClientLeft      =   2295
   ClientTop       =   2310
   ClientWidth     =   11250
   Icon            =   "OpeTra_frm_277.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6840
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   12065
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
         Height          =   2175
         Left            =   30
         TabIndex        =   21
         Top             =   4620
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   3836
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
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   0
            Left            =   1860
            TabIndex        =   4
            Top             =   60
            Width           =   3585
            _Version        =   65536
            _ExtentX        =   6324
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Juego de Planos completo del Inmueble"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   1
            Left            =   1860
            TabIndex        =   5
            Top             =   360
            Width           =   3405
            _Version        =   65536
            _ExtentX        =   6006
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Memoria Descriptiva"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   2
            Left            =   1860
            TabIndex        =   6
            Top             =   660
            Width           =   3345
            _Version        =   65536
            _ExtentX        =   5900
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Especificaciones Técnicas"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   3
            Left            =   1860
            TabIndex        =   7
            Top             =   960
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5636
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Lista de acabados"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   4
            Left            =   1860
            TabIndex        =   8
            Top             =   1260
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Presupuesto de Construcción"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   5
            Left            =   1860
            TabIndex        =   9
            Top             =   1560
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Estructura de Costos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   6
            Left            =   1860
            TabIndex        =   10
            Top             =   1860
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Licencia de Construcción"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   7
            Left            =   5820
            TabIndex        =   11
            Top             =   60
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Copia del Título de Propiedad"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   8
            Left            =   5820
            TabIndex        =   12
            Top             =   360
            Width           =   5265
            _Version        =   65536
            _ExtentX        =   9287
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "CRI completo o Copia Literal de Ficha Registral y Certif. de Gravamen"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   9
            Left            =   5820
            TabIndex        =   13
            Top             =   660
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "PU y HR del Terreno"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   10
            Left            =   5820
            TabIndex        =   14
            Top             =   960
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Copia de Declaratoria de Fábrica"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chk_Docume 
            Height          =   285
            Index           =   11
            Left            =   5820
            TabIndex        =   15
            Top             =   1260
            Width           =   4785
            _Version        =   65536
            _ExtentX        =   8440
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Copia de Escritura de Independización y Reglamento Interno"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label8 
            Caption         =   "Documentos a enviar:"
            Height          =   705
            Left            =   60
            TabIndex        =   22
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   435
         Left            =   30
         TabIndex        =   23
         Top             =   3060
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin VB.TextBox txt_ConObr 
            Height          =   315
            Left            =   1860
            MaxLength       =   60
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   60
            Width           =   9255
         End
         Begin VB.Label Label7 
            Caption         =   "Contacto Vendedor:"
            Height          =   285
            Left            =   60
            TabIndex        =   24
            Top             =   60
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1035
         Left            =   30
         TabIndex        =   25
         Top             =   3540
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   1826
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
         Begin VB.TextBox txt_Observ 
            Height          =   915
            Left            =   1860
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Text            =   "OpeTra_frm_277.frx":000C
            Top             =   60
            Width           =   9255
         End
         Begin VB.Label Label6 
            Caption         =   "Observaciones:"
            Height          =   285
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   27
         Top             =   2250
         Width           =   11145
         _Version        =   65536
         _ExtentX        =   19659
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
         Begin VB.ComboBox cmb_PerCon 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   9255
         End
         Begin VB.ComboBox cmb_EmpPer 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   9255
         End
         Begin VB.Label Label4 
            Caption         =   "Contacto:"
            Height          =   285
            Left            =   60
            TabIndex        =   29
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label5 
            Caption         =   "Empresa Peritaje:"
            Height          =   285
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   30
         Top             =   30
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
            Height          =   285
            Left            =   630
            TabIndex        =   31
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tasación del Inmueble"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   630
            TabIndex        =   32
            Top             =   330
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Emisión de Orden de Trabajo"
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
            Left            =   10560
            Top             =   90
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   9930
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   9360
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_277.frx":0010
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   33
         Top             =   1440
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1860
            TabIndex        =   34
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   9690
            TabIndex        =   35
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1860
            TabIndex        =   36
            Top             =   390
            Width           =   9255
            _Version        =   65536
            _ExtentX        =   16325
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
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   39
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8040
            TabIndex        =   38
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   37
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   40
         Top             =   750
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin VB.CommandButton cmd_EnvMail 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_277.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Envio de Correo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_277.frx":0624
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Imprimir Orden de Trabajo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "OpeTra_frm_277.frx":0A66
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_277.frx":0EA8
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_EvaTas_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_EmpPer()   As moddat_tpo_Genera
Dim l_str_Modali     As String
Dim l_dbl_ValVta     As Double
Dim l_str_TipVia     As String
Dim l_str_NomVia     As String
Dim l_str_NumVia     As String
Dim l_str_IntDpt     As String
Dim l_str_TipZon     As String
Dim l_str_NomZon     As String
Dim l_str_Refere     As String
Dim l_str_Depart     As String
Dim l_str_Provin     As String
Dim l_str_Distri     As String
Dim l_str_Estaci     As String
Dim l_str_DoiVen     As String
Dim l_str_NomVen     As String
Dim l_str_TelVen     As String
Dim l_str_Percon_mail As String

Private Sub chk_Docume_Click(Index As Integer, Value As Integer)
   cmd_Imprim.Enabled = False
   cmd_EnvMail.Enabled = False
End Sub

Private Sub cmb_EmpPer_Click()
   If cmb_EmpPer.ListIndex > -1 Then
      cmd_Imprim.Enabled = False
      cmd_EnvMail.Enabled = False
      
      Screen.MousePointer = 11
      Call moddat_gs_Carga_PerCon(cmb_PerCon, l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo, 1)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_EmpPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpPer_Click
   End If
End Sub

Private Sub cmb_PerCon_Click()
   cmd_Imprim.Enabled = False
   cmd_EnvMail.Enabled = False
   
   Call gs_SetFocus(txt_ConObr)
End Sub

Private Sub cmb_PerCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PerCon_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_Contad     As Integer
   Dim r_int_FlgDoc     As Integer
   
   If cmb_EmpPer.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa de Peritaje.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EmpPer)
      Exit Sub
   End If
   If cmb_PerCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Contacto en la Empresa de Peritaje.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerCon)
      Exit Sub
   End If
   If Len(Trim(txt_ConObr.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre del Contacto del Vendedor.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ConObr)
      Exit Sub
   End If
   
   r_int_FlgDoc = 1
   For r_int_Contad = 0 To 11
      If chk_Docume(r_int_Contad).Value = True Then
         r_int_FlgDoc = 2
      End If
   Next r_int_Contad
   
   If r_int_FlgDoc = 1 Then
      MsgBox "Debe seleccionar los Documentos a enviar.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_RPT_ORDTAS ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NomPrd & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Modali & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ValVta) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_TipVia & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_IntDpt & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_TipZon & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Estaci & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Distri & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Provin & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Depart & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_NomVen & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_DoiVen & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_TelVen & "', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'" & l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_PerCon.ItemData(cmb_PerCon.ListIndex), "000") & "', "
      g_str_Parame = g_str_Parame & "'" & txt_ConObr.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
      
      For r_int_Contad = 0 To 11
         If chk_Docume(r_int_Contad).Value = True Then
            g_str_Parame = g_str_Parame & "'1', "
         Else
            g_str_Parame = g_str_Parame & "'2', "
         End If
      Next r_int_Contad
      
      g_str_Parame = g_str_Parame & "'2', "
      g_str_Parame = g_str_Parame & "'2', "
      g_str_Parame = g_str_Parame & "'2', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 33, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 0
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct_1 = 2
   cmd_Imprim.Enabled = True
   cmd_EnvMail.Enabled = True
End Sub

Private Sub cmd_Imprim_Click()
   If MsgBox("¿Está seguro de Imprimir el reporte?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = "RPT_ORDTAS"
   crp_Imprim.DataFiles(1) = "CLI_DATGEN"
   crp_Imprim.DataFiles(2) = "MNT_PARDES"
   crp_Imprim.DataFiles(3) = "MNT_PERCON"
   
   crp_Imprim.SelectionFormula = "{RPT_ORDTAS.ORDTAS_NUMSOL} = '" & moddat_g_str_NumSol & "' AND {MNT_PARDES.PARDES_CODGRP} = '507' AND {MNT_PERCON.PERCON_TIPTAB} = 1 "
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_ORDTAS_11.RPT"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
   
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   moddat_g_int_FlgGrb = 1
   cmd_Imprim.Enabled = False
   cmd_EnvMail.Enabled = False

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM RPT_ORDTAS WHERE "
   g_str_Parame = g_str_Parame & "ORDTAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst

      cmb_EmpPer.ListIndex = gf_Busca_Arregl(l_arr_EmpPer, g_rst_Genera!ORDTAS_EMPPER) - 1
      Call moddat_gs_Carga_PerCon(cmb_PerCon, l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo, 1)
      Call gs_BuscarCombo_Item(cmb_PerCon, g_rst_Genera!ORDTAS_PERCON)
      txt_ConObr.Text = Trim(g_rst_Genera!ORDTAS_CONOBR & "")
      txt_Observ.Text = Trim(g_rst_Genera!ORDTAS_OBSERV & "")
      
      chk_Docume(0).Value = IIf(g_rst_Genera!ORDTAS_DOCR01 = "1", True, False)
      chk_Docume(1).Value = IIf(g_rst_Genera!ORDTAS_DOCR02 = "1", True, False)
      chk_Docume(2).Value = IIf(g_rst_Genera!ORDTAS_DOCR03 = "1", True, False)
      chk_Docume(3).Value = IIf(g_rst_Genera!ORDTAS_DOCR04 = "1", True, False)
      chk_Docume(4).Value = IIf(g_rst_Genera!ORDTAS_DOCR05 = "1", True, False)
      chk_Docume(5).Value = IIf(g_rst_Genera!ORDTAS_DOCR06 = "1", True, False)
      chk_Docume(6).Value = IIf(g_rst_Genera!ORDTAS_DOCR07 = "1", True, False)
      chk_Docume(7).Value = IIf(g_rst_Genera!ORDTAS_DOCR08 = "1", True, False)
      chk_Docume(8).Value = IIf(g_rst_Genera!ORDTAS_DOCR09 = "1", True, False)
      chk_Docume(9).Value = IIf(g_rst_Genera!ORDTAS_DOCR10 = "1", True, False)
      chk_Docume(10).Value = IIf(g_rst_Genera!ORDTAS_DOCR11 = "1", True, False)
      chk_Docume(11).Value = IIf(g_rst_Genera!ORDTAS_DOCR12 = "1", True, False)
   
      cmd_Imprim.Enabled = True
      cmd_EnvMail.Enabled = True
      moddat_g_int_FlgGrb = 2
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer
   
   'Cargando Lista de Empresas de Peritaje
   Call moddat_gs_Carga_LisIte(cmb_EmpPer, l_arr_EmpPer, 1, "507")

   cmb_PerCon.Clear
   txt_ConObr.Text = ""
   txt_Observ.Text = ""
   
   For r_int_Contad = 0 To 11
      chk_Docume(r_int_Contad).Value = False
   Next r_int_Contad
   
   'Obteniendo Datos del Cliente
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   l_str_Modali = ""
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLMAE_CODMOD)), "000")) Then
      l_str_Modali = moddat_g_arr_Genera(1).Genera_Nombre
   End If
   
   If g_rst_Princi!SOLMAE_TIPMON = 1 Then
      l_dbl_ValVta = g_rst_Princi!SOLMAE_COMVTA_SOL
   Else
      l_dbl_ValVta = g_rst_Princi!SOLMAE_COMVTA_DOL
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Obteniendo Datos del Inmueble
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      l_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA))
      l_str_NomVia = Trim(g_rst_Princi!SOLINM_NOMVIA & " ")
      l_str_NumVia = Trim(g_rst_Princi!SOLINM_NUMVIA & " ")
      l_str_IntDpt = Trim(g_rst_Princi!SOLINM_INTDPT & " ")
      l_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON))
      l_str_NomZon = Trim(g_rst_Princi!SOLINM_NOMZON & " ")
      l_str_Refere = Trim(g_rst_Princi!SOLINM_REFERE & "")
      l_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000")
      l_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00")
      l_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      l_str_Estaci = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      
      l_str_DoiVen = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
      If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Then
         l_str_NomVen = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
      Else
         l_str_NomVen = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO, l_str_TelVen)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_ConObr_Change()
   cmd_Imprim.Enabled = False
   cmd_EnvMail.Enabled = False
End Sub

Private Sub txt_ConObr_GotFocus()
   Call gs_SelecTodo(txt_ConObr)
End Sub

Private Sub txt_ConObr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Observ)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "()-_ .,;:¿?/&%$@#")
   End If
End Sub

Private Sub txt_Observ_Change()
   cmd_Imprim.Enabled = False
   cmd_EnvMail.Enabled = False
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(chk_Docume(0))
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub cmd_EnvMail_Click()
Dim r_str_NomPdf    As String
Dim r_str_Cadena    As String
Dim r_str_Parame    As String
Dim r_rst_Princi    As ADODB.Recordset

   moddat_g_str_CodGen = l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo
   frm_Tra_EvaTas_05.Show 1

   'Generando archivo PDF
'   r_str_NomPdf = ""
'   l_str_Percon_mail = ""
'   r_str_NomPdf = fs_GenExc
'
'   If Trim(l_str_Percon_mail) = "" Then
'      MsgBox "Empresa perito, el contacto no tiene ningun correo.", vbInformation, modgen_g_str_NomPlt
'   End If
'
'   If MsgBox("¿Está seguro de enviar el correo?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
'      Exit Sub
'   End If
'
'   If Trim(r_str_NomPdf) <> "" Then
'      'Enviando Correo Electrónico
'      modgen_g_str_Mail_Asunto = "ENVIO DE ORDEN DE TRABAJO (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
'      modgen_g_str_Mail_Mensaj = ""
'      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
'      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
'      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
'      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
'      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
'
'      'Seguimiento de instancias (Ocurrencias)
'      r_str_Parame = ""
'      r_str_Parame = r_str_Parame & " SELECT A.SEGDET_OBSERV, A.SEGFECCRE, A.SEGHORCRE "
'      r_str_Parame = r_str_Parame & "   FROM TRA_SEGDET A "
'      r_str_Parame = r_str_Parame & "  WHERE SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "'"
'      r_str_Parame = r_str_Parame & "    AND SEGDET_CODINS = 41 "
'      r_str_Parame = r_str_Parame & "    AND SEGDET_CODOCU = 21 "
'      r_str_Parame = r_str_Parame & "    AND SEGFECACT = 0 "
'      r_str_Parame = r_str_Parame & "  ORDER BY SEGDET_NUMOBS DESC "
'
'      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
'          Exit Sub
'      End If
'
'      If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
'         r_rst_Princi.MoveFirst
'         Do While Not r_rst_Princi.EOF
'            modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
'            modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "Fecha : " & gf_FormatoFecha(CStr(r_rst_Princi!SEGFECCRE)) & "  Hora : " & gf_FormatoHora(Format(r_rst_Princi!SEGHORCRE, "000000")) & Chr(13)
'            modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "DESCRIPCION OCURRENCIA : " & Trim(r_rst_Princi!SEGDET_OBSERV & "") & Chr(13)
'            r_rst_Princi.MoveNext
'         Loop
'      End If
'      r_rst_Princi.Close
'      Set r_rst_Princi = Nothing
'
'      'Adicionando aprobacion condicionada
'      r_str_Parame = ""
'      r_str_Parame = r_str_Parame & " SELECT TRIM(B.PARDES_DESCRI) AS INSTANCIA,  A.SEGCON_OBSCON, A.SEGCON_OBSLEV, A.SEGFECCRE, A.SEGHORCRE  "
'      r_str_Parame = r_str_Parame & "   FROM TRA_SEGCON A  "
'      r_str_Parame = r_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 2 AND B.PARDES_CODITE = A.SEGCON_CODINS  "
'      r_str_Parame = r_str_Parame & "  WHERE SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "'"
'      r_str_Parame = r_str_Parame & "    AND A.SEGCON_SITUAC = 1 "
'      r_str_Parame = r_str_Parame & "  ORDER BY SEGCON_SITUAC ASC, SEGCON_CODINS DESC  "
'
'      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
'          Exit Sub
'      End If
'
'      If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
'         r_rst_Princi.MoveFirst
'         Do While Not r_rst_Princi.EOF
'            If Trim(r_rst_Princi!SEGCON_OBSLEV & "") = "" Then
'               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
'               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "Fecha : " & gf_FormatoFecha(CStr(r_rst_Princi!SEGFECCRE)) & "  Hora : " & gf_FormatoHora(Format(r_rst_Princi!SEGHORCRE, "000000")) & Chr(13)
'               modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "CONDICION APROBACION : " & Trim(r_rst_Princi!SEGCON_OBSCON & "") & Chr(13)
'            End If
'            r_rst_Princi.MoveNext
'         Loop
'      End If
'      r_rst_Princi.Close
'      Set r_rst_Princi = Nothing
'
'      r_str_Cadena = l_str_Percon_mail
'      Call fs_Envia_CorreoOpe(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, "", "", "", 0, r_str_Cadena, r_str_NomPdf, g_str_RutLogTas)
'
'      r_str_Cadena = ""
'      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False, r_str_Cadena, r_str_NomPdf, g_str_RutLogTas)
'
'      'Creando Nueva Ocurrencia en Detalle de Seguimiento
'      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 93, 0, "", 0, 0) Then
'         Exit Sub
'      End If
'
'      MsgBox "Se envió la orden de trabajo por correo.", vbInformation, modgen_g_str_NomPlt
'
'      Screen.MousePointer = 0
'      moddat_g_int_FlgAct_1 = 2
'   End If
End Sub

Private Function fs_GenExc() As String
Dim r_rst_Princi      As ADODB.Recordset
Dim r_obj_Excel       As Excel.Application
Dim r_int_NumFil      As Integer
Dim r_str_Parame      As String

   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "  SELECT B.DATGEN_TIPDOC, B.DATGEN_NUMDOC, B.DATGEN_NUMCEL, B.DATGEN_TELEFO, A.ORDTAS_NUMSOL, "
   r_str_Parame = r_str_Parame & "         TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE) AS NOM_CLIENTE, "
   r_str_Parame = r_str_Parame & "         A.ORDTAS_PRODUC, A.ORDTAS_MODALI, A.ORDTAS_TIPMON, A.ORDTAS_VALVTA, "
   r_str_Parame = r_str_Parame & "         A.ORDTAS_TIPVIA, A.ORDTAS_NOMVIA, A.ORDTAS_NUMVIA, A.ORDTAS_INTDPT, "
   r_str_Parame = r_str_Parame & "         A.ORDTAS_TIPZON, A.ORDTAS_ESTACI, A.ORDTAS_DISTRI, A.ORDTAS_PROVIN, "
   r_str_Parame = r_str_Parame & "         A.ORDTAS_DEPART, A.ORDTAS_NOMVEN, A.ORDTAS_DOCVEN, A.ORDTAS_TELEF1, "
   r_str_Parame = r_str_Parame & "         A.ORDTAS_CONOBR, A.ORDTAS_OBSERV, A.ORDTAS_DOCR01, A.ORDTAS_DOCR02, "
   r_str_Parame = r_str_Parame & "         A.ORDTAS_DOCR03, A.ORDTAS_DOCR04, A.ORDTAS_DOCR05, A.ORDTAS_DOCR06, "
   r_str_Parame = r_str_Parame & "         A.ORDTAS_DOCR07, A.ORDTAS_DOCR08, A.ORDTAS_DOCR09, A.ORDTAS_DOCR10, "
   r_str_Parame = r_str_Parame & "         A.ORDTAS_DOCR11, A.ORDTAS_DOCR12, A.ORDTAS_NOMZON, A.ORDTAS_REFERE, "
   r_str_Parame = r_str_Parame & "         C.PARDES_DESCRI AS EMPR_PERITAJE, D.PERCON_NOMBRE AS CONTACTO, d.percon_direle "
   r_str_Parame = r_str_Parame & "    FROM RPT_ORDTAS A "
   r_str_Parame = r_str_Parame & "   INNER JOIN CLI_DATGEN B ON A.ORDTAS_TDOCLI = B.DATGEN_TIPDOC AND A.ORDTAS_NDOCLI = B.DATGEN_NUMDOC "
   r_str_Parame = r_str_Parame & "   INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 507 AND C.PARDES_CODITE = A.ORDTAS_EMPPER "
   r_str_Parame = r_str_Parame & "   INNER JOIN MNT_PERCON D ON D.PERCON_CODEMP = A.ORDTAS_EMPPER AND D.PERCON_CODCON = A.ORDTAS_PERCON AND D.PERCON_TIPTAB = 1 "
   r_str_Parame = r_str_Parame & "   WHERE A.ORDTAS_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Screen.MousePointer = 0
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Function
   End If
   
   r_rst_Princi.MoveFirst
   l_str_Percon_mail = Trim(r_rst_Princi!PERCON_DIRELE & "")
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      'MARGENES
      .PageSetup.LeftMargin = Application.CentimetersToPoints(1.5)
      .PageSetup.RightMargin = Application.CentimetersToPoints(0.4)
      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
      .Columns("A").ColumnWidth = 13
      .Columns("D").ColumnWidth = 2
      .Columns("G").ColumnWidth = 10
      .Columns("H").ColumnWidth = 5
      .Columns("F").ColumnWidth = 10
      .Columns("I").ColumnWidth = 10
      .Range(.Cells(1, 1), .Cells(67, 12)).Font.Name = "Arial (Western)"
      .Range(.Cells(1, 1), .Cells(67, 12)).Font.Size = 8
      .Range(.Cells(1, 1), .Cells(67, 12)).RowHeight = 14
      
      .Pictures.Insert(g_str_RutLog & "\" & "image001.gif").Select
      
      .Cells(7, 1) = "ORDEN DE TRABAJO - TASACION DE INMUEBLE"
      .Range(.Cells(7, 1), .Cells(7, 9)).Merge
      .Range(.Cells(7, 1), .Cells(7, 9)).Font.Bold = True
      .Range(.Cells(7, 1), .Cells(7, 9)).Font.Underline = True
      .Range(.Cells(7, 1), .Cells(7, 9)).HorizontalAlignment = xlHAlignCenter

      .Rows(1).RowHeight = 1
      .Rows(8).RowHeight = 9
      .Rows(9).RowHeight = 5
      .Range(.Cells(9, 1), .Cells(9, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      '.CELDA(FILA, COLUMNA)
      .Cells(2, 7) = "Nombre Reporte:"
      .Cells(3, 7) = "Fecha Emisión:"
      .Cells(4, 7) = "Hora Emisión:"
      .Cells(5, 7) = "Página:"
      
      .Cells(2, 9) = "OPE_ORDTAS_11"
      .Cells(3, 9) = Format(date, "dd/mm/yyyy")
      .Cells(4, 9) = Format(Time, "hh:mm:ss")
      .Cells(5, 9) = "1"
      .Range(.Cells(2, 9), .Cells(5, 9)).HorizontalAlignment = xlHAlignRight
      
      .Cells(10, 1) = "Nro Solicitud:"
      .Cells(10, 2) = gf_Formato_NumSol(moddat_g_str_NumSol)
      .Range(.Cells(10, 1), .Cells(10, 10)).Font.Bold = True
      .Rows(11).RowHeight = 5
      .Rows(12).RowHeight = 5
      .Range(.Cells(12, 1), .Cells(12, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Cells(13, 1) = "Empresa Peritaje:"
      .Cells(13, 2) = Trim(r_rst_Princi!EMPR_PERITAJE & "")
      .Cells(14, 2) = Trim(r_rst_Princi!CONTACTO & "")
      
      .Rows(15).RowHeight = 5
      .Rows(16).RowHeight = 5
      .Range(.Cells(16, 1), .Cells(16, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Cells(17, 1) = "Producto:"
      .Cells(18, 1) = "Moneda"
      .Cells(18, 6) = "Modalidad:"
      .Cells(19, 1) = "Valor Venta:"
      .Cells(20, 1) = "Cliente:"
      .Cells(21, 1) = "DOI Cliente": .Cells(21, 6) = "Telefono Fijo:": .Cells(21, 8) = "Celular:"

      .Range(.Cells(18, 7), .Cells(20, 9)).Merge
      .Range(.Cells(18, 7), .Cells(20, 9)).WrapText = True
      .Range(.Cells(18, 7), .Cells(20, 9)).VerticalAlignment = xlTop

      .Cells(17, 2) = Trim(r_rst_Princi!ORDTAS_PRODUC & "")
      .Cells(18, 2) = IIf(r_rst_Princi!ORDTAS_TIPMON = 1, "NUEVOS SOLES", "DOLARES AMERICANOS")
      .Cells(18, 7) = Trim(r_rst_Princi!ORDTAS_MODALI & "")
      .Cells(19, 2) = IIf(r_rst_Princi!ORDTAS_TIPMON = 1, "S/.", "US$") & " " & Format(r_rst_Princi!ORDTAS_VALVTA, "###,###,###,##0.00")
      .Cells(20, 2) = Trim(r_rst_Princi!NOM_CLIENTE & "")
      .Cells(21, 2) = IIf(r_rst_Princi!DatGen_TipDoc = 1, "DNI", "CE") & " - " & Trim(r_rst_Princi!DATGEN_NUMDOC)
      .Cells(21, 7) = Trim(r_rst_Princi!DatGen_Telefo & "")
      .Cells(21, 9) = Trim(r_rst_Princi!DATGEN_NUMCEL & "")

      .Rows(22).RowHeight = 5
      .Rows(23).RowHeight = 5
      .Range(.Cells(23, 1), .Cells(23, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Cells(24, 1) = "Tipo de Vía:": .Cells(24, 6) = "Nombre Vía:"
      .Cells(25, 1) = "Número:"
      .Cells(26, 1) = "Int/Dpt/Mz/Lt.:"
      .Cells(27, 1) = "Tipo de Zona:": .Cells(27, 6) = "Nombre Zona:"
      .Cells(28, 1) = "Estacionamiento:"
      .Cells(29, 1) = "Distrito:"
      .Cells(30, 1) = "Provincia:"
      .Cells(31, 1) = "Departamento:"
      .Cells(32, 1) = "Referencia:"
                  
      .Range(.Cells(24, 7), .Cells(26, 9)).Merge
      .Range(.Cells(24, 7), .Cells(26, 9)).WrapText = True
      .Range(.Cells(24, 7), .Cells(26, 9)).VerticalAlignment = xlTop
      
      .Cells(24, 2) = Trim(r_rst_Princi!ORDTAS_TIPVIA & "")
      .Cells(24, 7) = Trim(r_rst_Princi!ORDTAS_NOMVIA & "")
      .Cells(25, 2) = Trim(r_rst_Princi!ORDTAS_NUMVIA & "")
      .Cells(26, 2) = Trim(r_rst_Princi!ORDTAS_INTDPT & "")
      .Cells(27, 2) = Trim(r_rst_Princi!ORDTAS_TIPZON & "")
      .Cells(27, 7) = Trim(r_rst_Princi!ORDTAS_NOMZON & "")
      .Cells(28, 2) = Trim(r_rst_Princi!ORDTAS_ESTACI & "")
      .Cells(29, 2) = Trim(r_rst_Princi!ORDTAS_DISTRI & "")
      .Cells(30, 2) = Trim(r_rst_Princi!ORDTAS_PROVIN & "")
      .Cells(31, 2) = Trim(r_rst_Princi!ORDTAS_DEPART & "")
      .Cells(32, 2) = Trim(r_rst_Princi!ORDTAS_REFERE & "")
      
      .Rows(33).RowHeight = 5
      .Rows(34).RowHeight = 5
      .Range(.Cells(34, 1), .Cells(34, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous

      .Cells(35, 1) = "Vendedor:"
      .Cells(36, 1) = "DOI Vendedor:": .Cells(36, 6) = "Teléfono:"
      
      .Cells(35, 2) = Trim(r_rst_Princi!ORDTAS_NOMVEN & "")
      .Cells(36, 2) = Trim(r_rst_Princi!ORDTAS_DOCVEN & "")
      .Cells(36, 7) = Trim(r_rst_Princi!ORDTAS_TELEF1 & "")
      
      .Rows(37).RowHeight = 5
      .Rows(38).RowHeight = 5
      .Range(.Cells(38, 1), .Cells(38, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous

      .Cells(39, 1) = "Sirvase realizar el informe de Tasación (Original y Copia) del Inmueble con los Datos adjuntos, para lo cual hacemos entrega de"
      .Cells(40, 1) = "los siguientes documento:"

      .Cells(42, 1) = IIf(r_rst_Princi!ORDTAS_DOCR01 = "1", "[ X ] ", "[   ] ") & "Juego de Planos completos del Inmueble":
      .Cells(43, 1) = IIf(r_rst_Princi!ORDTAS_DOCR02 = "1", "[ X ] ", "[   ] ") & "Memoria Descriptiva"
      .Cells(44, 1) = IIf(r_rst_Princi!ORDTAS_DOCR03 = "1", "[ X ] ", "[   ] ") & "Especificaciones Técnicas"
      .Cells(45, 1) = IIf(r_rst_Princi!ORDTAS_DOCR04 = "1", "[ X ] ", "[   ] ") & "Lista de acabados"
      .Cells(46, 1) = IIf(r_rst_Princi!ORDTAS_DOCR05 = "1", "[ X ] ", "[   ] ") & "Presupuesto de Construcción"
      .Cells(47, 1) = IIf(r_rst_Princi!ORDTAS_DOCR06 = "1", "[ X ] ", "[   ] ") & "Estructura de Costos"
      .Cells(48, 1) = IIf(r_rst_Princi!ORDTAS_DOCR07 = "1", "[ X ] ", "[   ] ") & "Licencia de Contrucción"

      .Cells(42, 5) = IIf(r_rst_Princi!ORDTAS_DOCR08 = "1", "[ X ] ", "[   ] ") & "Copia del TÍtulo de Propiedad inscrito en RRPP o RPU":
      .Cells(43, 5) = IIf(r_rst_Princi!ORDTAS_DOCR09 = "1", "[ X ] ", "[   ] ") & "CRI completo (RRPP) o Copia Literal de la Ficha Registral y"
      .Cells(44, 5) = "      Certificado de Gravamen (RPU) del Terreno"
      .Cells(45, 5) = IIf(r_rst_Princi!ORDTAS_DOCR10 = "1", "[ X ] ", "[   ] ") & "PU y HR del Terreno"
      .Cells(46, 5) = IIf(r_rst_Princi!ORDTAS_DOCR11 = "1", "[ X ] ", "[   ] ") & "Copia de Declaratoria de Fábrica"
      .Cells(47, 5) = IIf(r_rst_Princi!ORDTAS_DOCR12 = "1", "[ X ] ", "[   ] ") & "Copia de la Escritura de Independización y Reglamento Interno"
      
      .Rows(49).RowHeight = 5
      .Rows(50).RowHeight = 5
      .Range(.Cells(50, 1), .Cells(50, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous

      .Cells(51, 1) = "Persona Contacto:"
      .Cells(52, 1) = "Observaciones:"
      
      .Cells(51, 2) = Trim(r_rst_Princi!ORDTAS_CONOBR & "")
      .Cells(52, 2) = Trim(r_rst_Princi!ORDTAS_OBSERV & "")
            
      .Range(.Cells(54, 1), .Cells(54, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
            
      .Range(.Cells(60, 1), .Cells(60, 3)).Merge
      .Cells(60, 1) = "miCasita hipotecaria"
      .Cells(60, 1).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(60, 1), .Cells(60, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Range(.Cells(52, 2), .Cells(53, 9)).Merge
      .Range(.Cells(52, 2), .Cells(53, 9)).WrapText = True
      .Range(.Cells(52, 2), .Cells(53, 9)).VerticalAlignment = xlTop

   End With
   
   'r_obj_Excel.ActiveWorkbook.SaveAs (l_str_rutarz & "\" & l_str_Cadena & ".XLSX")
   fs_GenExc = ""
   fs_GenExc = moddat_g_str_NumSol & "_41_1_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".PDF"
   
   r_obj_Excel.ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, FileName:=g_str_RutLogTas & fs_GenExc, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
   r_obj_Excel.ActiveWorkbook.Close SaveChanges:=False
   
   r_obj_Excel.Application.Quit
   Set r_obj_Excel = Nothing
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   Screen.MousePointer = 0
End Function

