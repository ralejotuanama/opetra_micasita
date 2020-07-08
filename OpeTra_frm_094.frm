VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_EvaTas_12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6870
   ClientLeft      =   7725
   ClientTop       =   3120
   ClientWidth     =   11250
   Icon            =   "OpeTra_frm_094.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6840
      Left            =   0
      TabIndex        =   19
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
         TabIndex        =   38
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
            TabIndex        =   39
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   435
         Left            =   30
         TabIndex        =   36
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
            TabIndex        =   37
            Top             =   60
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1035
         Left            =   30
         TabIndex        =   34
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
            Text            =   "OpeTra_frm_094.frx":000C
            Top             =   60
            Width           =   9255
         End
         Begin VB.Label Label6 
            Caption         =   "Observaciones:"
            Height          =   285
            Left            =   60
            TabIndex        =   35
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   31
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
         Begin VB.ComboBox cmb_EmpPer 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   9255
         End
         Begin VB.ComboBox cmb_PerCon 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   9255
         End
         Begin VB.Label Label5 
            Caption         =   "Empresa Peritaje:"
            Height          =   285
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "Contacto:"
            Height          =   285
            Left            =   60
            TabIndex        =   32
            Top             =   390
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   20
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
            TabIndex        =   21
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
            TabIndex        =   30
            Top             =   330
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Orden de Trabajo"
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_094.frx":0010
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   22
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
            TabIndex        =   23
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   9690
            TabIndex        =   24
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1860
            TabIndex        =   25
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
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   28
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8040
            TabIndex        =   27
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   29
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_094.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "OpeTra_frm_094.frx":075C
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_094.frx":0B9E
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Imprimir Orden de Trabajo"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_EvaTas_12"
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

Private Sub chk_Docume_Click(Index As Integer, Value As Integer)
   cmd_Imprim.Enabled = False
End Sub

Private Sub cmb_EmpPer_Click()
   If cmb_EmpPer.ListIndex > -1 Then
      cmd_Imprim.Enabled = False
      
      Screen.MousePointer = 11
      Call moddat_gs_Carga_PerCon(cmb_PerCon, l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo)
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
   
   moddat_g_int_FlgAct_1 = 2
   
   cmd_Imprim.Enabled = True
End Sub

Private Sub cmd_Imprim_Click()
   If MsgBox("¿Está seguro de Imprimir el reporte?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".RPT_ORDTAS"
   crp_Imprim.DataFiles(1) = UCase(moddat_g_str_EntDat) & ".CLI_DATGEN"
   crp_Imprim.DataFiles(2) = UCase(moddat_g_str_EntDat) & ".MNT_PARDES"
   crp_Imprim.DataFiles(3) = UCase(moddat_g_str_EntDat) & ".MNT_PERCON"
   
   crp_Imprim.SelectionFormula = "{RPT_ORDTAS.ORDTAS_NUMSOL} = '" & moddat_g_str_NumSol & "' AND {MNT_PARDES.PARDES_CODGRP} = '507' "
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_ORDTAS_11.RPT"
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

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM RPT_ORDTAS WHERE "
   g_str_Parame = g_str_Parame & "ORDTAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst

      cmb_EmpPer.ListIndex = gf_Busca_Arregl(l_arr_EmpPer, g_rst_Genera!ORDTAS_EMPPER) - 1
      
      Call moddat_gs_Carga_PerCon(cmb_PerCon, l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo)
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
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_ConObr_Change()
   cmd_Imprim.Enabled = False
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
