VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ges_CreHip_18 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   5850
   ClientTop       =   3945
   ClientWidth     =   11250
   Icon            =   "OpeTra_frm_154.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7320
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   12912
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
         Height          =   2085
         Left            =   30
         TabIndex        =   6
         Top             =   5190
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin VB.TextBox txt_NumPol_Viv 
            Height          =   315
            Left            =   1860
            MaxLength       =   120
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   1710
            Width           =   2925
         End
         Begin Threed.SSPanel pnl_FecEva_Viv 
            Height          =   315
            Left            =   1860
            TabIndex        =   7
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "01/01/1999"
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
         Begin Threed.SSPanel pnl_TipApl_Viv 
            Height          =   315
            Left            =   1860
            TabIndex        =   8
            Top             =   720
            Width           =   2715
            _Version        =   65536
            _ExtentX        =   4789
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "FACTOR"
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
         Begin Threed.SSPanel pnl_ValApl_Viv 
            Height          =   315
            Left            =   1860
            TabIndex        =   9
            Top             =   1050
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.02"
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
         Begin Threed.SSPanel pnl_FecEmi_Viv 
            Height          =   315
            Left            =   1860
            TabIndex        =   36
            Top             =   1380
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "18/10/2000"
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
         Begin VB.Label Label15 
            Caption         =   "F. Emisión:"
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label14 
            Caption         =   "Seguro Inmueble"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label Label16 
            Caption         =   "Valor Aplicación:"
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   1050
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "Tipo Aplicación:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label18 
            Caption         =   "F. Evaluación:"
            Height          =   285
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label11 
            Caption         =   "Nro. de Póliza:"
            Height          =   285
            Left            =   60
            TabIndex        =   10
            Top             =   1710
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2415
         Left            =   30
         TabIndex        =   16
         Top             =   2730
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin VB.TextBox txt_NumPoC_Des 
            Height          =   315
            Left            =   4800
            MaxLength       =   120
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   2040
            Width           =   2925
         End
         Begin VB.TextBox txt_NumPol_Des 
            Height          =   315
            Left            =   1860
            MaxLength       =   120
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   2040
            Width           =   2925
         End
         Begin Threed.SSPanel pnl_TipSeg 
            Height          =   315
            Left            =   1860
            TabIndex        =   17
            Top             =   390
            Width           =   9255
            _Version        =   65536
            _ExtentX        =   16325
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "INDIVIDUAL"
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
         Begin Threed.SSPanel pnl_FecEva_Des 
            Height          =   315
            Left            =   1860
            TabIndex        =   18
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "01/01/1999"
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
         Begin Threed.SSPanel pnl_TipApl_Des 
            Height          =   315
            Left            =   1860
            TabIndex        =   19
            Top             =   1050
            Width           =   2715
            _Version        =   65536
            _ExtentX        =   4789
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "FACTOR"
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
         Begin Threed.SSPanel pnl_ValApl_Des 
            Height          =   315
            Left            =   1860
            TabIndex        =   20
            Top             =   1380
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.02"
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
         Begin Threed.SSPanel pnl_FecEmi_Des 
            Height          =   315
            Left            =   1860
            TabIndex        =   35
            Top             =   1710
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "18/10/2000"
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
         Begin VB.Label Label25 
            Caption         =   "Nro. de Póliza:"
            Height          =   285
            Left            =   60
            TabIndex        =   27
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "F. Emisión:"
            Height          =   285
            Left            =   60
            TabIndex        =   26
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label Label12 
            Caption         =   "F. Evaluación:"
            Height          =   285
            Left            =   60
            TabIndex        =   25
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo Seguro Desgrav.:"
            Height          =   285
            Left            =   60
            TabIndex        =   24
            Top             =   390
            Width           =   1665
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo Aplicación:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Valor Aplicación:"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   1380
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Seguro Desgravamen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   3315
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   28
         Top             =   2250
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
         Begin Threed.SSPanel pnl_EmpSeg 
            Height          =   315
            Left            =   1860
            TabIndex        =   29
            Top             =   60
            Width           =   9255
            _Version        =   65536
            _ExtentX        =   16325
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
         Begin VB.Label Label5 
            Caption         =   "Empresa Seguros:"
            Height          =   285
            Left            =   60
            TabIndex        =   30
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   31
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
            Height          =   315
            Left            =   660
            TabIndex        =   33
            Top             =   30
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   315
            Left            =   660
            TabIndex        =   34
            Top             =   330
            Width           =   6885
            _Version        =   65536
            _ExtentX        =   12144
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Modificación de Números de Pólizas de Seguros"
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
            Picture         =   "OpeTra_frm_154.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   32
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
         Begin VB.CommandButton cmd_ImpMulti 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_154.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Imprimir Seguro para Vivienda"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpCertif 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_154.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Imprimir Seguro Desgravamen"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_154.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "OpeTra_frm_154.frx":0FDC
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   37
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1860
            TabIndex        =   38
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1860
            TabIndex        =   39
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   41
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   40
            Top             =   60
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_CreHip_18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_NumPol_Des.Text)) = 0 Then
      MsgBox "Debe ingresar el Nro. de Póliza.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumPol_Des)
      Exit Sub
   End If

   If Len(Trim(txt_NumPol_Viv.Text)) = 0 Then
      MsgBox "Debe ingresar el Nro. de Póliza.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumPol_Viv)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_POLIZA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumPol_Des.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumPoC_Des.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumPol_Viv.Text & "', "
      g_str_Parame = g_str_Parame & Format(CDate(pnl_FecEmi_Des.Caption), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(CDate(pnl_FecEmi_Viv.Caption), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
      
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
   Loop
   
   Screen.MousePointer = 0
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_ImpCertif_Click()
Dim i As Integer

   g_str_Parame = "" '
   g_str_Parame = g_str_Parame & " SELECT B.DATGEN_TIPDOC DATGEN_TIPODOCC, B.DATGEN_NUMDOC DATGEN_NUMDOCC, B.DATGEN_FLGDOA FLGDOAC, B.DATGEN_REGCYG DATGEN_REGCYGC,"
   g_str_Parame = g_str_Parame & "        B.DATGEN_TIPDOA DATGEN_TIPDOAC, B.DATGEN_NUMDOA DATGEN_NUMDOAC, B.DATGEN_NACPAI DATGEN_NACPAIC,"
   g_str_Parame = g_str_Parame & "        B.DATGEN_NOMBRE DATGEN_NOMBREC, B.DATGEN_APEPAT DATGEN_APEPATC, B.DATGEN_APEMAT DATGEN_APEMATC,"
   g_str_Parame = g_str_Parame & "        B.DATGEN_NUMDOC DATGEN_NUMDOCC, B.DATGEN_NACFEC DATGEN_NACFECC, B.DATGEN_CODSEX DATGEN_CODSEXC,"
   g_str_Parame = g_str_Parame & "        B.DATGEN_ESTCIV DATGEN_ESTCIVC, B.DATGEN_PROFES DATGEN_PROFESC, B.DATGEN_TIPVIA DATGEN_TIPVIAC,"
   g_str_Parame = g_str_Parame & "        B.DATGEN_NOMVIA DATGEN_NOMVIAC, B.DATGEN_NUMERO DATGEN_NUMEROC, B.DATGEN_INTDPT DATGEN_INTDPTC,"
   g_str_Parame = g_str_Parame & "        B.DATGEN_NOMZON DATGEN_NOMZONC, B.DATGEN_TIPZON DATGEN_TIPZONC, B.DATGEN_UBIGEO DATGEN_UBIGEOC,"
   g_str_Parame = g_str_Parame & "        B.DATGEN_CYGNDO DATGEN_CYGNDOC, B.DATGEN_NIVEST DATGEN_NIVESTC, B.DATGEN_NUMCEL DATGEN_NUMCELC,"
   g_str_Parame = g_str_Parame & "        B.DATGEN_DIRELE DATGEN_DIRELEC, B.DATGEN_REFERE DATGEN_REFEREC, B.DATGEN_TELEFO DATGEN_TELEFOC,"
   g_str_Parame = g_str_Parame & "        C.DATGEN_TIPDOC DATGEN_TIPODOCY, C.DATGEN_NUMDOC DATGEN_NUMDOCY, C.DATGEN_FLGDOA FLGDOAY, C.DATGEN_REGCYG DATGEN_REGCYGY,"
   g_str_Parame = g_str_Parame & "        C.DATGEN_TIPDOA DATGEN_TIPDOAY, C.DATGEN_NUMDOA DATGEN_NUMDOAY, C.DATGEN_NACPAI DATGEN_NACPAIY,"
   g_str_Parame = g_str_Parame & "        C.DATGEN_NOMBRE DATGEN_NOMBREY, C.DATGEN_APEPAT DATGEN_APEPATY, C.DATGEN_APEMAT DATGEN_APEMATY,"
   g_str_Parame = g_str_Parame & "        C.DATGEN_NUMDOC DATGEN_NUMDOCY, C.DATGEN_NACFEC DATGEN_NACFECY, C.DATGEN_CODSEX DATGEN_CODSEXY,"
   g_str_Parame = g_str_Parame & "        C.DATGEN_ESTCIV DATGEN_ESTCIVY, C.DATGEN_PROFES DATGEN_PROFESY, C.DATGEN_TIPVIA DATGEN_TIPVIAY,"
   g_str_Parame = g_str_Parame & "        C.DATGEN_NOMVIA DATGEN_NOMVIAY, C.DATGEN_NUMERO DATGEN_NUMEROY, C.DATGEN_INTDPT DATGEN_INTDPTY,"
   g_str_Parame = g_str_Parame & "        C.DATGEN_NOMZON DATGEN_NOMZONY, C.DATGEN_TIPZON DATGEN_TIPZONY, C.DATGEN_UBIGEO DATGEN_UBIGEOY,"
   g_str_Parame = g_str_Parame & "        C.DATGEN_CYGNDO DATGEN_CYGNDOY, C.DATGEN_NIVEST DATGEN_NIVESTY, C.DATGEN_NUMCEL DATGEN_NUMCELY,"
   g_str_Parame = g_str_Parame & "        C.DATGEN_DIRELE DATGEN_DIRELEY, C.DATGEN_REFERE DATGEN_REFEREY, C.DATGEN_TELEFO DATGEN_TELEFOY,"
   g_str_Parame = g_str_Parame & "        HIPMAE_MONEDA, HIPMAE_SEGPRE, HIPMAE_TIPSEG, HIPMAE_MTOPRE, HIPMAE_PLAANO,"
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(A1.PARDES_DESCRI) FROM MNT_PARDES A1 WHERE A1.PARDES_CODGRP = '101' AND A1.PARDES_CODITE = SUBSTR(B.DATGEN_UBIGEO,1,2)||'0000') AS DEPARTAMENTO_C,"
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(A2.PARDES_DESCRI) FROM MNT_PARDES A2 WHERE A2.PARDES_CODGRP = '101' AND A2.PARDES_CODITE = SUBSTR(B.DATGEN_UBIGEO,1,4)||'00') AS PROVINCIA_C,"
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(A3.PARDES_DESCRI) FROM MNT_PARDES A3 WHERE A3.PARDES_CODGRP = '101' AND A3.PARDES_CODITE = B.DATGEN_UBIGEO) AS DISTRITO_C,"
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(A1.PARDES_DESCRI) FROM MNT_PARDES A1 WHERE A1.PARDES_CODGRP = '101' AND A1.PARDES_CODITE = SUBSTR(C.DATGEN_UBIGEO,1,2)||'0000') AS DEPARTAMENTO_Y,"
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(A2.PARDES_DESCRI) FROM MNT_PARDES A2 WHERE A2.PARDES_CODGRP = '101' AND A2.PARDES_CODITE = SUBSTR(C.DATGEN_UBIGEO,1,4)||'00') AS PROVINCIA_Y,"
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(A3.PARDES_DESCRI) FROM MNT_PARDES A3 WHERE A3.PARDES_CODGRP = '101' AND A3.PARDES_CODITE = C.DATGEN_UBIGEO) AS DISTRITO_Y"
           
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A"
   g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN B ON A.HIPMAE_NDOCLI=B.DATGEN_NUMDOC"
   g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN C ON A.HIPMAE_NDOCYG=C.DATGEN_NUMDOC"
   g_str_Parame = g_str_Parame & "  WHERE HIPMAE_NUMOPE='" & moddat_g_str_NumOpe & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      For i = 1 To 10
         Printer.FontSize = 10
         Printer.Print ""
      Next
      Printer.FontSize = 2
      Printer.Print ""
      Printer.Print ""
      Printer.FontSize = 10
      With g_rst_Princi
      Printer.Print Tab(4); Trim(!DATGEN_NOMBREC) & " " & Trim(!DATGEN_APEPATC) & " " & Trim(!DATGEN_APEMATC); Tab(105); Mid(!DATGEN_NACFECC, 7, 2) & "    " & Mid(!DATGEN_NACFECC, 5, 2) & "    " & Mid(!DATGEN_NACFECC, 1, 4)
      For i = 1 To 1
          Printer.FontSize = 10
          Printer.Print ""
      Next
      Printer.FontSize = 3
      Printer.Print ""
      Printer.Print ""
      Printer.FontSize = 10
      
      Printer.Print Tab(4); IIf(!DATGEN_TIPODOCC = 1, "X", "       X"); Tab(30); Trim(!DATGEN_NUMDOCC & ""); Tab(65); IIf(!DATGEN_CODSEXC = 1, " X", "      X"); Tab(73); moddat_gf_Consulta_ParDes("500", Trim(!DATGEN_NACPAIC)); _
                    Tab(99); IIf(!DATGEN_ESTCIVC = 1, "X", IIf(!DATGEN_ESTCIVC = 2, "     X", IIf(!DATGEN_ESTCIVC = 3, "          X", "               X")))
                       
      Printer.FontSize = 10
      Printer.Print ""
      Printer.FontSize = 1
      Printer.Print ""
      Printer.FontSize = 10
      
      Printer.Print Tab(4); Mid(moddat_gf_Consulta_ParDes("201", CStr(!DATGEN_TIPVIAC)) & _
                               " " & Trim(!DATGEN_NOMVIAC) & " " & Trim(!DATGEN_NUMEROC) & _
                               IIf(Len(Trim(!DATGEN_INTDPTC)) > 0, " (" & Trim(!DATGEN_INTDPTC) & ")", "") & _
                               IIf(Len(Trim(!DATGEN_NOMZONC)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(!DATGEN_TIPZONC)) & " " & Trim(!DATGEN_NOMZONC), ""), 1, 82) _
                               ; Tab(105); Mid(Trim(!DISTRITO_C), 1, 15);
      For i = 1 To 2
          Printer.FontSize = 10
          Printer.Print ""
      Next
      Printer.FontSize = 2
      Printer.Print ""
      Printer.FontSize = 10
      Printer.Print Tab(5); Trim(!PROVINCIA_C) & "/" & Trim(!DEPARTAMENTO_C); Tab(48); Trim(!DATGEN_DIRELEC); Tab(106); Trim(!DATGEN_TELEFOC)
      For i = 1 To 2
          Printer.FontSize = 7
          Printer.Print ""
      Next
      Printer.FontSize = 10
      Printer.Print Tab(5); moddat_gf_Consulta_ParDes("501", CStr(!DATGEN_PROFESC));
      For i = 1 To 4
          Printer.FontSize = 10
          Printer.Print ""
      Next
      Printer.FontSize = 2
      Printer.Print ""
      Printer.FontSize = 10
      If moddat_gf_Consulta_TipSeg(!HIPMAE_SEGPRE, !HIPMAE_TIPSEG) = "DESG. MANCOMUNADO" Then
         Printer.Print Tab(5); Trim(!DATGEN_NOMBREY) & " " & Trim(!DATGEN_APEPATY) & " " & Trim(!DATGEN_APEMATY); Tab(90); Trim(!DATGEN_NUMDOCY & "");
            
         For i = 1 To 2
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.Print Tab(5); Mid(moddat_gf_Consulta_ParDes("201", CStr(!DATGEN_TIPVIAC)) & _
                               " " & Trim(!DATGEN_NOMVIAC) & " " & Trim(!DATGEN_NUMEROC) & _
                               IIf(Len(Trim(!DATGEN_INTDPTC)) > 0, " (" & Trim(!DATGEN_INTDPTC) & ")", "") & _
                               IIf(Len(Trim(!DATGEN_NOMZONC)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(!DATGEN_TIPZONC)) & " " & Trim(!DATGEN_NOMZONC), ""), 1, 82) _
                               ; Tab(105); Mid(Trim(!DISTRITO_C), 1, 15);
         For i = 1 To 2
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.FontSize = 3
         Printer.Print ""
         Printer.FontSize = 10
         Printer.Print Tab(5); Trim(!PROVINCIA_Y) & "/" & Trim(!DEPARTAMENTO_Y); Tab(50); !DATGEN_TELEFOC; Tab(68); Trim("") 'PARENTESCO
      End If

         Printer.EndDoc
      End With
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   '*******************************************
'   Dim i As Integer
'
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT B.DATGEN_TIPDOC DATGEN_TIPODOCC, B.DATGEN_NUMDOC DATGEN_NUMDOCC, B.DATGEN_FLGDOA FLGDOAC, B.DATGEN_REGCYG DATGEN_REGCYGC,"
'   g_str_Parame = g_str_Parame & "        B.DATGEN_TIPDOA DATGEN_TIPDOAC, B.DATGEN_NUMDOA DATGEN_NUMDOAC, B.DATGEN_NACPAI DATGEN_NACPAIC,"
'   g_str_Parame = g_str_Parame & "        B.DATGEN_NOMBRE DATGEN_NOMBREC, B.DATGEN_APEPAT DATGEN_APEPATC, B.DATGEN_APEMAT DATGEN_APEMATC,"
'   g_str_Parame = g_str_Parame & "        B.DATGEN_NUMDOC DATGEN_NUMDOCC, B.DATGEN_NACFEC DATGEN_NACFECC, B.DATGEN_CODSEX DATGEN_CODSEXC,"
'   g_str_Parame = g_str_Parame & "        B.DATGEN_ESTCIV DATGEN_ESTCIVC, B.DATGEN_PROFES DATGEN_PROFESC, B.DATGEN_TIPVIA DATGEN_TIPVIAC,"
'   g_str_Parame = g_str_Parame & "        B.DATGEN_NOMVIA DATGEN_NOMVIAC, B.DATGEN_NUMERO DATGEN_NUMEROC, B.DATGEN_INTDPT DATGEN_INTDPTC,"
'   g_str_Parame = g_str_Parame & "        B.DATGEN_NOMZON DATGEN_NOMZONC, B.DATGEN_TIPZON DATGEN_TIPZONC, B.DATGEN_UBIGEO DATGEN_UBIGEOC,"
'   g_str_Parame = g_str_Parame & "        B.DATGEN_CYGNDO DATGEN_CYGNDOC, B.DATGEN_NIVEST DATGEN_NIVESTC, B.DATGEN_NUMCEL DATGEN_NUMCELC,"
'   g_str_Parame = g_str_Parame & "        B.DATGEN_DIRELE DATGEN_DIRELEC, B.DATGEN_REFERE DATGEN_REFEREC, B.DATGEN_TELEFO DATGEN_TELEFOC,"
'   g_str_Parame = g_str_Parame & "        C.DATGEN_TIPDOC DATGEN_TIPODOCY, C.DATGEN_NUMDOC DATGEN_NUMDOCY, C.DATGEN_FLGDOA FLGDOAY, C.DATGEN_REGCYG DATGEN_REGCYGY,"
'   g_str_Parame = g_str_Parame & "        C.DATGEN_TIPDOA DATGEN_TIPDOAY, C.DATGEN_NUMDOA DATGEN_NUMDOAY, C.DATGEN_NACPAI DATGEN_NACPAIY,"
'   g_str_Parame = g_str_Parame & "        C.DATGEN_NOMBRE DATGEN_NOMBREY, C.DATGEN_APEPAT DATGEN_APEPATY, C.DATGEN_APEMAT DATGEN_APEMATY,"
'   g_str_Parame = g_str_Parame & "        C.DATGEN_NUMDOC DATGEN_NUMDOCY, C.DATGEN_NACFEC DATGEN_NACFECY, C.DATGEN_CODSEX DATGEN_CODSEXY,"
'   g_str_Parame = g_str_Parame & "        C.DATGEN_ESTCIV DATGEN_ESTCIVY, C.DATGEN_PROFES DATGEN_PROFESY, C.DATGEN_TIPVIA DATGEN_TIPVIAY,"
'   g_str_Parame = g_str_Parame & "        C.DATGEN_NOMVIA DATGEN_NOMVIAY, C.DATGEN_NUMERO DATGEN_NUMEROY, C.DATGEN_INTDPT DATGEN_INTDPTY,"
'   g_str_Parame = g_str_Parame & "        C.DATGEN_NOMZON DATGEN_NOMZONY, C.DATGEN_TIPZON DATGEN_TIPZONY, C.DATGEN_UBIGEO DATGEN_UBIGEOY,"
'   g_str_Parame = g_str_Parame & "        C.DATGEN_CYGNDO DATGEN_CYGNDOY, C.DATGEN_NIVEST DATGEN_NIVESTY, C.DATGEN_NUMCEL DATGEN_NUMCELY,"
'   g_str_Parame = g_str_Parame & "        C.DATGEN_DIRELE DATGEN_DIRELEY, C.DATGEN_REFERE DATGEN_REFEREY, C.DATGEN_TELEFO DATGEN_TELEFOY,"
'   g_str_Parame = g_str_Parame & "        HIPMAE_MONEDA, HIPMAE_SEGPRE, HIPMAE_TIPSEG, HIPMAE_MTOPRE, HIPMAE_PLAANO"
'   g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A"
'   g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN B ON A.HIPMAE_NDOCLI=B.DATGEN_NUMDOC"
'   g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN C ON A.HIPMAE_NDOCYG=C.DATGEN_NUMDOC"
'   g_str_Parame = g_str_Parame & "  WHERE HIPMAE_NUMOPE='" & moddat_g_str_NumOpe & "'"
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'       Exit Sub
'   End If
'
'   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
'      For i = 1 To 12
'         Printer.FontSize = 10
'         Printer.Print ""
'      Next
'
'      With g_rst_Princi
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Printer.Print Tab(0); Trim(!DATGEN_NOMBREC); Tab(56); Trim(!DATGEN_APEPATC); Tab(102); Trim(!DATGEN_APEMATC)
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'
'         Select Case moddat_gf_Consulta_ParDes("203", CStr(!DATGEN_TIPODOCC))
'            Case "DNI"
'               Printer.Print Tab(4); "X"; Tab(33); Trim(!DATGEN_NUMDOCC & ""); Tab(80); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECC)), 1, 2); Tab(87); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECC)), 4, 2); Tab(94); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECC)), 7); Tab(113); IIf(Mid(moddat_gf_Consulta_ParDes("207", CStr(!DATGEN_CODSEXC)), 1, 1) = "F", "X", ""); Tab(117); IIf(Mid(moddat_gf_Consulta_ParDes("207", CStr(!DATGEN_CODSEXC)), 1, 1) = "M", "X", "")
'            Case "CARNE DE EXTRANJERIA"
'               Printer.Print Tab(10); "X"; Tab(33); Trim(!DATGEN_NUMDOCC & ""); Tab(80); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECC)), 1, 2); Tab(87); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECC)), 4, 2); Tab(94); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECC)), 7); Tab(113); IIf(Mid(moddat_gf_Consulta_ParDes("207", CStr(!DATGEN_CODSEXC)), 1, 1) = "F", "X", ""); Tab(117); IIf(Mid(moddat_gf_Consulta_ParDes("207", CStr(!DATGEN_CODSEXC)), 1, 1) = "M", "X", "")
'         End Select
'
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Printer.Print Tab(2); Trim(moddat_gf_Consulta_ParDes("205", CStr(!DATGEN_ESTCIVC)) & IIf(!DATGEN_ESTCIVC = 2, " / " & moddat_gf_Consulta_ParDes("206", !DATGEN_REGCYGC), "")); Tab(72); Mid(moddat_gf_Consulta_ParDes("501", CStr(!DATGEN_PROFESC)), 1, 37)
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Printer.Print Tab(0); moddat_gf_Consulta_ParDes("201", CStr(!DATGEN_TIPVIAC)) & _
'                               " " & Trim(!DATGEN_NOMVIAC) & " " & Trim(!DATGEN_NUMEROC) & _
'                               IIf(Len(Trim(!DATGEN_INTDPTC)) > 0, " (" & Trim(!DATGEN_INTDPTC) & ")", "") & _
'                               IIf(Len(Trim(!DATGEN_NOMZONC)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(!DATGEN_TIPZONC)) & " " & Trim(!DATGEN_NOMZONC), "")
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Printer.Print Tab(0); Mid(moddat_gf_Consulta_ParDes("101", Trim(!DATGEN_UBIGEOC)), 1, 22); Tab(47); moddat_gf_Consulta_ParDes("101", Left(!DATGEN_UBIGEOC, 4) & "00"); Tab(99); moddat_gf_Consulta_ParDes("101", Left(!DATGEN_UBIGEOC, 2) & "0000")
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Printer.Print Tab(1); Mid(Trim(!DATGEN_REFEREC & ""), 1, 31); Tab(103); moddat_gf_Consulta_ParDes("500", Trim(!DATGEN_NACPAIC))
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Printer.Print Tab(1); Trim(!DATGEN_TELEFOC & ""); Tab(40); Trim(!DATGEN_NUMCELC & ""); Tab(70); Mid(Trim(!DATGEN_DIRELEC & ""), 1, 33)
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Printer.Print ""
'
'         If moddat_gf_Consulta_TipSeg(!HIPMAE_SEGPRE, !HIPMAE_TIPSEG) = "DESG. MANCOMUNADO" Then
'
'            Printer.FontSize = 1: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(0); Trim(!DATGEN_NOMBREY); Tab(56); Trim(!DATGEN_APEPATY); Tab(102); Trim(!DATGEN_APEMATY)
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'
'            Select Case moddat_gf_Consulta_ParDes("203", CStr(!DATGEN_TIPODOCY))
'               Case "DNI"
'                  Printer.Print Tab(4); "X"; Tab(33); Trim(!DATGEN_NUMDOCY & ""); Tab(80); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECY)), 1, 2); Tab(87); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECY)), 4, 2); Tab(94); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECY)), 7); Tab(113); IIf(Mid(moddat_gf_Consulta_ParDes("207", CStr(!DATGEN_CODSEXY)), 1, 1) = "F", "X", ""); Tab(117); IIf(Mid(moddat_gf_Consulta_ParDes("207", CStr(!DATGEN_CODSEXY)), 1, 1) = "M", "X", "")
'               Case "CARNE DE EXTRANJERIA"
'                  Printer.Print Tab(10); "X"; Tab(33); Trim(!DATGEN_NUMDOCY & ""); Tab(80); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECY)), 1, 2); Tab(87); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECY)), 4, 2); Tab(94); Mid(gf_FormatoFecha(CStr(!DATGEN_NACFECY)), 7); Tab(113); IIf(Mid(moddat_gf_Consulta_ParDes("207", CStr(!DATGEN_CODSEXY)), 1, 1) = "F", "X", ""); Tab(117); IIf(Mid(moddat_gf_Consulta_ParDes("207", CStr(!DATGEN_CODSEXY)), 1, 1) = "M", "X", "")
'            End Select
'
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(2); Trim(moddat_gf_Consulta_ParDes("205", CStr(!DATGEN_ESTCIVY)) & IIf(!DATGEN_ESTCIVY = 2, " / " & moddat_gf_Consulta_ParDes("206", !DATGEN_REGCYGY), "")); Tab(72); Mid(moddat_gf_Consulta_ParDes("501", CStr(!DATGEN_PROFESY)), 1, 37)
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(0); moddat_gf_Consulta_ParDes("201", CStr(!DATGEN_TIPVIAC)) & _
'                                  " " & Trim(!DATGEN_NOMVIAC) & " " & Trim(!DATGEN_NUMEROC) & _
'                                  IIf(Len(Trim(!DATGEN_INTDPTC)) > 0, " (" & Trim(!DATGEN_INTDPTC) & ")", "") & _
'                                  IIf(Len(Trim(!DATGEN_NOMZONC)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(!DATGEN_TIPZONC)) & " " & Trim(!DATGEN_NOMZONC), "")
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(0); Mid(moddat_gf_Consulta_ParDes("101", Trim(!DATGEN_UBIGEOC)), 1, 22); Tab(47); moddat_gf_Consulta_ParDes("101", Left(!DATGEN_UBIGEOC, 4) & "00"); Tab(99); moddat_gf_Consulta_ParDes("101", Left(!DATGEN_UBIGEOC, 2) & "0000")
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(1); Mid(Trim(!DATGEN_REFEREC & ""), 1, 31); Tab(103); moddat_gf_Consulta_ParDes("500", Trim(!DATGEN_NACPAIY))
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(1); Trim(!DATGEN_TELEFOC & ""); Tab(40); Trim(!DATGEN_NUMCELY & ""); Tab(70); Mid(Trim(!DATGEN_DIRELEY & ""), 1, 33)
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print ""
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Else
'            Printer.FontSize = 1: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(0); ""
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(4); ""
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(0); ""
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(0); ""
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(0); ""
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(1); ""
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print Tab(0); ""
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'            Printer.Print ""
'            Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         End If
'
'         If moddat_gf_Consulta_ParDes("204", CStr(!HIPMAE_MONEDA)) = "SOLES" Then
'            Printer.Print Tab(0); "X"; Tab(60); Format(!HIPMAE_MTOPRE, "#,###,##0.00"); Tab(103); "HIPOTECARIO"
'         Else
'            Printer.Print Tab(14); "X"; Tab(60); Format(!HIPMAE_MTOPRE, "#,###,##0.00"); Tab(103); "HIPOTECARIO"
'         End If
'         Printer.FontSize = 2: Printer.Print "": Printer.Print "": Printer.Print ""
'         Printer.FontSize = 10
'
'         If moddat_gf_Consulta_TipSeg(!HIPMAE_SEGPRE, !HIPMAE_TIPSEG) = "DESG. INDIVIDUAL" Then
'            Printer.Print Tab(30); "X"; Tab(41); Format(!HIPMAE_PLAANO); Tab(70); "X"
'         Else
'            Printer.Print Tab(30); "X"; Tab(41); Format(!HIPMAE_PLAANO); Tab(89); "X"
'         End If
'
'         Printer.EndDoc
'      End With
'   End If
'
'   g_rst_Princi.Close
'   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_ImpMulti_Click()
Dim i As Integer
Dim strNomCli As String

   g_str_Parame = "" '
   g_str_Parame = g_str_Parame & " SELECT A.HIPMAE_NUMOPE AS OPERACION,"
   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_APEPAT) DATGEN_APEPAT, TRIM(B.DATGEN_APEMAT) DATGEN_APEMAT, TRIM(B.DATGEN_APECAS) DATGEN_APECAS, TRIM(B.DATGEN_NOMBRE) DATGEN_NOMBRE,"
   g_str_Parame = g_str_Parame & "        SUBSTR(B.DATGEN_NACFEC,7,2)||'    '||SUBSTR(B.DATGEN_NACFEC,5,2)||'    '||SUBSTR(B.DATGEN_NACFEC,1,4) AS FECHA_NACIMIENTO, "
   g_str_Parame = g_str_Parame & "        B.DATGEN_TIPDOC AS TIPO_DOC, TRIM(B.DATGEN_NUMDOC) AS DNI, TRIM(H.PARDES_DESCRI) AS NACIONALIDAD,"
   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_NUMCEL)||'-'||TRIM(B.DATGEN_TELEFO) AS TELEFONO,"
   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_DIRELE) AS CORREO,"
   g_str_Parame = g_str_Parame & "        SUBSTR(A.HIPMAE_FECDES,7,2)||'/'||SUBSTR(A.HIPMAE_FECDES,5,2)||'/'||SUBSTR(A.HIPMAE_FECDES,1,4) AS FECHA_DESEMBOLSO,"
   g_str_Parame = g_str_Parame & "        SUBSTR(A.HIPMAE_ULTVCT,7,2)||'/'||SUBSTR(A.HIPMAE_ULTVCT,5,2)||'/'||SUBSTR(A.HIPMAE_ULTVCT,1,4) AS ULTIMA_CUOTA,"
   g_str_Parame = g_str_Parame & "        C.EVATAS_SUMASE_INM+C.EVATAS_SUMASE_ES1+C.EVATAS_SUMASE_ES2+C.EVATAS_SUMASE_DEP AS SUMA_ASEGURADA,"
   g_str_Parame = g_str_Parame & "        TRIM(D.PARDES_DESCRI) AS CARACT_BIEN_ASEGURADO,"
   g_str_Parame = g_str_Parame & "        1 AS USO_INMUEBLE, C.EVATAS_TIPMON, "
   g_str_Parame = g_str_Parame & "        C.EVATAS_NUMSOT AS NUMERO_SOTANOS,"
   g_str_Parame = g_str_Parame & "        TRIM(E.PARDES_CODITE) AS TIPO_CONSTRUCCION,"
   g_str_Parame = g_str_Parame & "        C.EVATAS_ANOCON AS ANIO_CONSTRUCCION, "
   g_str_Parame = g_str_Parame & "        C.EVATAS_NUMPIS AS NUMERO_PISOS, "
   g_str_Parame = g_str_Parame & "        3 AS MATERIAL_CONSTRUCCION,"
   g_str_Parame = g_str_Parame & "        DATGEN_TIPVIA,DATGEN_NOMVIA,DATGEN_NUMERO,DATGEN_INTDPT,DATGEN_NOMZON,DATGEN_TIPZON,DATGEN_NOMZON,DATGEN_UBIGEO,"
   g_str_Parame = g_str_Parame & "        SOLINM_TIPVIA,SOLINM_NOMVIA,SOLINM_NUMVIA,SOLINM_INTDPT,SOLINM_INTDPT,SOLINM_NOMZON,SOLINM_TIPZON,SOLINM_UBIGEO,EVALEG_FEENIN,"
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(A1.PARDES_DESCRI) FROM MNT_PARDES A1 WHERE A1.PARDES_CODGRP = '101' AND A1.PARDES_CODITE = SUBSTR(B.DATGEN_UBIGEO,1,2)||'0000') AS DEPARTAMENTO,"
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(A2.PARDES_DESCRI) FROM MNT_PARDES A2 WHERE A2.PARDES_CODGRP = '101' AND A2.PARDES_CODITE = SUBSTR(B.DATGEN_UBIGEO,1,4)||'00') AS PROVINCIA,"
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(A3.PARDES_DESCRI) FROM MNT_PARDES A3 WHERE A3.PARDES_CODGRP = '101' AND A3.PARDES_CODITE = B.DATGEN_UBIGEO) AS DISTRITO,"
   g_str_Parame = g_str_Parame & "        (SELECT TRIM(A4.PARDES_DESCRI) FROM MNT_PARDES A4 WHERE A4.PARDES_CODGRP = '501' AND A4.PARDES_CODITE = B.DATGEN_PROFES) AS PROFESION,"
   
   g_str_Parame = g_str_Parame & "        CASE WHEN EVATAS_MATCON = 1 THEN '                                                                X                     X '"
   g_str_Parame = g_str_Parame & "             WHEN EVATAS_MATCON = 5 THEN ' X '"
   g_str_Parame = g_str_Parame & "             WHEN EVATAS_MATCON = 3 THEN '       X '"
   g_str_Parame = g_str_Parame & "        END As MATERIAL_CONSTRUCCION, "
   
   g_str_Parame = g_str_Parame & "       (CASE WHEN LENGTH(TRIM(F.SOLINM_INTDPT)) > 0 THEN"
   g_str_Parame = g_str_Parame & "                  (SELECT TRIM(B1.PARDES_DESCRI) FROM MNT_PARDES B1 WHERE PARDES_CODGRP = '201' AND PARDES_CODITE ="
   g_str_Parame = g_str_Parame & "                          LPAD(TRIM(F.SOLINM_TIPVIA),6,0)) || ' ' || TRIM(F.SOLINM_NOMVIA) || ' ' ||"
   g_str_Parame = g_str_Parame & "                          TRIM(F.SOLINM_NUMVIA) || ' - DPTO / INT.: ' || TRIM(F.SOLINM_INTDPT)"
   g_str_Parame = g_str_Parame & "        ELSE (SELECT TRIM(B1.PARDES_DESCRI) FROM MNT_PARDES B1 WHERE PARDES_CODGRP = '201' AND PARDES_CODITE ="
   g_str_Parame = g_str_Parame & "                     LPAD(TRIM(F.SOLINM_TIPVIA),6,0))"
   g_str_Parame = g_str_Parame & "                     || ' ' || TRIM(F.SOLINM_NOMVIA) || ' ' || TRIM(F.SOLINM_NUMVIA)"
   g_str_Parame = g_str_Parame & "        END ||"
   g_str_Parame = g_str_Parame & "        CASE WHEN LENGTH(TRIM(F.SOLINM_NOMZON)) > 0 THEN"
   g_str_Parame = g_str_Parame & "                  ' - '|| (SELECT TRIM(B1.PARDES_DESCRI) FROM MNT_PARDES B1 WHERE PARDES_CODGRP = '202' AND PARDES_CODITE ="
   g_str_Parame = g_str_Parame & "                  LPAD(TRIM(F.SOLINM_TIPZON),6,0))||' '||TRIM(F.SOLINM_NOMZON)"
   g_str_Parame = g_str_Parame & "        END) AS DIRECCION_INMUEBLE,"
           
   g_str_Parame = g_str_Parame & "       (SELECT TRIM(A1.PARDES_DESCRI) FROM MNT_PARDES A1 WHERE A1.PARDES_CODGRP = '101' AND A1.PARDES_CODITE = SUBSTR(F.SOLINM_UBIGEO,1,2)||'0000') AS DEPARTAMENTO_INMUEBLE,"
   g_str_Parame = g_str_Parame & "       (SELECT TRIM(A2.PARDES_DESCRI) FROM MNT_PARDES A2 WHERE A2.PARDES_CODGRP = '101' AND A2.PARDES_CODITE = SUBSTR(F.SOLINM_UBIGEO,1,4)||'00') AS PROVINCIA_INMUEBLE,"
   g_str_Parame = g_str_Parame & "       (SELECT TRIM(A3.PARDES_DESCRI) FROM MNT_PARDES A3 WHERE A3.PARDES_CODGRP = '101' AND A3.PARDES_CODITE = F.SOLINM_UBIGEO) AS DISTRITO_INMUEBLE"
           
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A"
   g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.HIPMAE_TDOCLI AND B.DATGEN_NUMDOC = A.HIPMAE_NDOCLI"
   g_str_Parame = g_str_Parame & "  INNER JOIN TRA_EVATAS C ON C.EVATAS_NUMSOL = A.HIPMAE_NUMSOL"
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 223 AND D.PARDES_CODITE = C.EVATAS_MATCON"
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 221 AND E.PARDES_CODITE = C.EVATAS_TIPINM"
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_SOLINM F ON F.SOLINM_NUMSOL = C.EVATAS_NUMSOL"
   g_str_Parame = g_str_Parame & "  INNER JOIN TRA_EVALEG G ON A.HIPMAE_NUMSOL = G.EVALEG_NUMSOL" 'añadido 11/04/2016
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES H ON H.PARDES_CODGRP = '500' AND H.PARDES_CODITE = B.DATGEN_PAIRES "
   g_str_Parame = g_str_Parame & "  WHERE A.HIPMAE_NUMOPE='" & moddat_g_str_NumOpe & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
      
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      For i = 1 To 16
         Printer.FontSize = 10
         Printer.Print ""
      Next
      Printer.FontSize = 3
      Printer.Print ""
      Printer.Print ""
      Printer.FontSize = 10
      With g_rst_Princi
         If Not IsNull(!DATGEN_APECAS) Then
            If Len(Trim(!DATGEN_APECAS)) > 0 Then
               strNomCli = Trim(!DatGen_ApePat) & " " & Trim(!DatGen_ApeMat) & " DE " & Trim(!DATGEN_APECAS) & " " & Trim(!DatGen_Nombre)
            Else
               strNomCli = Trim(!DatGen_ApePat) & " " & Trim(!DatGen_ApeMat) & " " & Trim(!DatGen_Nombre)
            End If
         Else
            strNomCli = Trim(!DatGen_ApePat) & " " & Trim(!DatGen_ApeMat) & " " & Trim(!DatGen_Nombre)
         End If
         
         Printer.Print Tab(4); "  " & strNomCli; Tab(102); "  " & !FECHA_NACIMIENTO;
         For i = 1 To 2
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.FontSize = 2
         Printer.Print ""
         Printer.FontSize = 10
         Printer.Print Tab(4); IIf(!TIPO_DOC = 1, "   X", IIf(!TIPO_DOC = 2, "         X", "                     X")); Tab(29); Trim(!DNI); Tab(67); !NACIONALIDAD; Tab(102); " " & Trim(!TELEFONO);
         For i = 1 To 2
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.Print Tab(4); "  " & Mid(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                               " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                               IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & _
                               " " & Trim(g_rst_Princi!DatGen_NomZon), "") & " " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo)), 1, 68); Tab(102); " " & Mid(Trim(!DISTRITO), 1, 15);
         For i = 1 To 2
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.FontSize = 2
         Printer.Print ""
         Printer.FontSize = 10
         Printer.Print Tab(4); "  " & Trim(!PROVINCIA) & "/" & Trim(!DEPARTAMENTO); Tab(67); IIf(IsNull(!CORREO), "", !CORREO);
         For i = 1 To 2
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.Print Tab(4); ""; Tab(67); Mid(Trim(!PROFESION), 1, 40);
         For i = 1 To 4
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.FontSize = 4
         Printer.Print ""
         Printer.FontSize = 10
         Printer.Print Tab(39); Trim(!ANIO_CONSTRUCCION); Tab(75); Trim(!NUMERO_PISOS); Tab(109); Trim(!NUMERO_SOTANOS)
         For i = 1 To 1
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.Print Tab(4); !MATERIAL_CONSTRUCCION
         For i = 1 To 1
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.FontSize = 2
         Printer.Print ""
         Printer.FontSize = 10
         Printer.Print Tab(4); "  " & Mid(Trim(!DIRECCION_INMUEBLE), 1, 68); Tab(100); "" & Mid(!DISTRITO_INMUEBLE, 1, 15)
         For i = 1 To 1
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.FontSize = 3
         Printer.Print ""
         Printer.FontSize = 10
         Printer.Print Tab(4); "  " & Trim(!PROVINCIA_INMUEBLE) & "/" & Trim(!DEPARTAMENTO_INMUEBLE);
         '*****************************************
         For i = 1 To 22
             Printer.FontSize = 10
             Printer.Print ""
         Next
         Printer.FontSize = 3
         Printer.Print ""
         Printer.FontSize = 10
         Printer.Print Tab(26); IIf(!EVATAS_TIPMON = 1, "S/. ", "US$ ") & Format(!SUMA_ASEGURADA, "#,###,##0.00")

         Printer.EndDoc
      End With
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   '*************************************************************
'   Dim i As Integer
'   Dim strNomCli As String
'
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT A.HIPMAE_NUMOPE AS OPERACION,"
'   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_APEPAT) DATGEN_APEPAT, TRIM(B.DATGEN_APEMAT) DATGEN_APEMAT, TRIM(B.DATGEN_APECAS) DATGEN_APECAS, TRIM(B.DATGEN_NOMBRE) DATGEN_NOMBRE,"
'   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_NUMDOC) AS DNI,"
'   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_NUMCEL)||'   '||TRIM(B.DATGEN_TELEFO) AS TELEFONO,"
'   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_DIRELE) AS CORREO,"
'   g_str_Parame = g_str_Parame & "        SUBSTR(A.HIPMAE_FECDES,7,2)||'/'||SUBSTR(A.HIPMAE_FECDES,5,2)||'/'||SUBSTR(A.HIPMAE_FECDES,1,4) AS FECHA_DESEMBOLSO,"
'   g_str_Parame = g_str_Parame & "        SUBSTR(A.HIPMAE_ULTVCT,7,2)||'/'||SUBSTR(A.HIPMAE_ULTVCT,5,2)||'/'||SUBSTR(A.HIPMAE_ULTVCT,1,4) AS ULTIMA_CUOTA,"
'   g_str_Parame = g_str_Parame & "        C.EVATAS_SUMASE_INM+C.EVATAS_SUMASE_ES1+C.EVATAS_SUMASE_ES2+C.EVATAS_SUMASE_DEP AS SUMA_ASEGURADA,"
'   g_str_Parame = g_str_Parame & "        TRIM(D.PARDES_DESCRI) AS CARACT_BIEN_ASEGURADO,"
'   g_str_Parame = g_str_Parame & "        1 AS USO_INMUEBLE,"
'   g_str_Parame = g_str_Parame & "        C.EVATAS_NUMSOT AS NUMERO_SOTANOS,"
'   g_str_Parame = g_str_Parame & "        TRIM(E.PARDES_CODITE) AS TIPO_CONSTRUCCION,"
'   g_str_Parame = g_str_Parame & "        C.EVATAS_ANOCON AS ANIO_CONSTRUCCION, "
'   g_str_Parame = g_str_Parame & "        C.EVATAS_NUMPIS AS NUMERO_PISOS, "
'   g_str_Parame = g_str_Parame & "        3 AS MATERIAL_CONSTRUCCION,"
'   g_str_Parame = g_str_Parame & "        DATGEN_TIPVIA,DATGEN_NOMVIA,DATGEN_NUMERO,DATGEN_INTDPT,DATGEN_NOMZON,DATGEN_TIPZON,DATGEN_NOMZON,DATGEN_UBIGEO,"
'   g_str_Parame = g_str_Parame & "        SOLINM_TIPVIA,SOLINM_NOMVIA,SOLINM_NUMVIA,SOLINM_INTDPT,SOLINM_INTDPT,SOLINM_NOMZON,SOLINM_TIPZON,SOLINM_UBIGEO,EVALEG_FEENIN"
'   g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A"
'   g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.HIPMAE_TDOCLI AND B.DATGEN_NUMDOC = A.HIPMAE_NDOCLI"
'   g_str_Parame = g_str_Parame & "  INNER JOIN TRA_EVATAS C ON C.EVATAS_NUMSOL = A.HIPMAE_NUMSOL"
'   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 223 AND D.PARDES_CODITE = C.EVATAS_MATCON"
'   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 221 AND E.PARDES_CODITE = C.EVATAS_TIPINM"
'   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_SOLINM F ON F.SOLINM_NUMSOL = C.EVATAS_NUMSOL"
'   g_str_Parame = g_str_Parame & "  INNER JOIN TRA_EVALEG G ON A.HIPMAE_NUMSOL = G.EVALEG_NUMSOL" 'añadido 11/04/2016
'   g_str_Parame = g_str_Parame & "  WHERE A.HIPMAE_NUMOPE='" & moddat_g_str_NumOpe & "'"
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'       Exit Sub
'   End If
'
'   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
'      For i = 1 To 19 '22
'         Printer.FontSize = 10
'         Printer.Print ""
'      Next
'
'      With g_rst_Princi
'         If Not IsNull(!DATGEN_APECAS) Then
'            If Len(Trim(!DATGEN_APECAS)) > 0 Then
'               strNomCli = Trim(!DatGen_ApePat) & " " & Trim(!DatGen_ApeMat) & " DE " & Trim(!DATGEN_APECAS) & " " & Trim(!DatGen_Nombre)
'            Else
'               strNomCli = Trim(!DatGen_ApePat) & " " & Trim(!DatGen_ApeMat) & " " & Trim(!DatGen_Nombre)
'            End If
'         Else
'            strNomCli = Trim(!DatGen_ApePat) & " " & Trim(!DatGen_ApeMat) & " " & Trim(!DatGen_Nombre)
'         End If
'
'         Printer.Print Tab(30); strNomCli; Tab(110); Trim(!DNI)
'
'         'utilizamos esta lineacon tipo de fuente 4 para imprimir solo en blanco con este y luego a su tamaño normal(10)
'         Printer.FontSize = 4: Printer.Print "":  Printer.FontSize = 10
'         Printer.Print Tab(5); Mid(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
'                                        " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
'                                        IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
'                                        IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "") & " " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo)), 1, 80)
'
'         'de igual manera utilizamos esta linea y la imprimimos con este tamaño(4) y luego al tamaño normal(10)
'         'para poder hacer ingresar Direccion como un tamaño normal que se encuentra entre ambas lineas en blanco,
'         'debido a que cada vez hacia que imprima en una linea mas ancha y nos descuadraba el formato del documento
'         Printer.FontSize = 4: Printer.Print "": Printer.Print "":  Printer.FontSize = 10
'
'         Printer.Print Tab(5); Trim(!TELEFONO); Tab(65); IIf(IsNull(!CORREO), "", !CORREO)
'         Printer.Print ""
'
'         If frm_Ges_CreHip_02.grd_Listad.TextMatrix(7, 1) = "BIEN TERMINADO" Then
'            Printer.Print Tab(40); Trim(!FECHA_DESEMBOLSO); Tab(90); Trim(!ULTIMA_CUOTA)
'         Else
'            If IsNull(!EVALEG_FEENIN) Then
'               Printer.Print Tab(40); !FECHA_DESEMBOLSO; Tab(90); Trim(!ULTIMA_CUOTA)
'            Else
'               If (!EVALEG_FEENIN) = 0 Then
'                  Printer.Print Tab(40); !FECHA_DESEMBOLSO; Tab(90); Trim(!ULTIMA_CUOTA)
'               Else
'                  Printer.Print Tab(40); gf_FormatoFecha(CStr(!EVALEG_FEENIN)); Tab(90); Trim(!ULTIMA_CUOTA)
'               End If
'            End If
'         End If
'
'         For i = 1 To 3
'            Printer.Print ""
'         Next
'
'         Printer.Print Tab(12); Format(!SUMA_ASEGURADA, "#,###,##0.00")
'         Printer.FontSize = 4: Printer.Print "": Printer.Print "": Printer.FontSize = 10
'         Printer.Print Tab(12); Mid(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA), 1, 35) & " " & Trim(g_rst_Princi!SOLINM_INTDPT); Tab(90); moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Printer.Print Tab(12); moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00")
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Printer.Print Tab(12); Trim(!NUMERO_PISOS); Tab(80); Trim(!NUMERO_SOTANOS)
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'         Printer.Print Tab(20); Trim(!ANIO_CONSTRUCCION)
'         Printer.FontSize = 4: Printer.Print "":  Printer.FontSize = 10
'         Printer.Print Tab(30); Trim(!CARACT_BIEN_ASEGURADO)
'
'         For i = 1 To 7
'            Printer.Print ""
'         Next
'         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
'
'         Printer.EndDoc
'      End With
'   End If
'
'   g_rst_Princi.Close
'   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_ImpMulti_Click_old()
Dim i As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.HIPMAE_NUMOPE AS OPERACION,"
   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE) AS NOMBRE_COMPLETO,"
   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_NUMDOC) AS DNI,"
   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_NUMCEL)||'   '||TRIM(B.DATGEN_TELEFO) AS TELEFONO,"
   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_DIRELE) AS CORREO,"
   g_str_Parame = g_str_Parame & "        SUBSTR(A.HIPMAE_FECDES,7,2)||'/'||SUBSTR(A.HIPMAE_FECDES,5,2)||'/'||SUBSTR(A.HIPMAE_FECDES,1,4) AS FECHA_DESEMBOLSO,"
   g_str_Parame = g_str_Parame & "        SUBSTR(A.HIPMAE_ULTVCT,7,2)||'/'||SUBSTR(A.HIPMAE_ULTVCT,5,2)||'/'||SUBSTR(A.HIPMAE_ULTVCT,1,4) AS ULTIMA_CUOTA,"
   g_str_Parame = g_str_Parame & "        C.EVATAS_SUMASE_INM+C.EVATAS_SUMASE_ES1+C.EVATAS_SUMASE_ES2+C.EVATAS_SUMASE_DEP AS SUMA_ASEGURADA,"
   g_str_Parame = g_str_Parame & "        TRIM(D.PARDES_DESCRI) AS CARACT_BIEN_ASEGURADO,"
   g_str_Parame = g_str_Parame & "        1 AS USO_INMUEBLE,"
   g_str_Parame = g_str_Parame & "        C.EVATAS_NUMSOT AS NUMERO_SOTANOS,"
   g_str_Parame = g_str_Parame & "        TRIM(E.PARDES_CODITE) AS TIPO_CONSTRUCCION,"
   g_str_Parame = g_str_Parame & "        C.EVATAS_ANOCON AS ANIO_CONSTRUCCION, "
   g_str_Parame = g_str_Parame & "        C.EVATAS_NUMPIS AS NUMERO_PISOS, "
   g_str_Parame = g_str_Parame & "        3 AS MATERIAL_CONSTRUCCION,"
   g_str_Parame = g_str_Parame & "        DATGEN_TIPVIA,DATGEN_NOMVIA,DATGEN_NUMERO,DATGEN_INTDPT,DATGEN_NOMZON,DATGEN_TIPZON,DATGEN_NOMZON,DATGEN_UBIGEO,"
   g_str_Parame = g_str_Parame & "        SOLINM_TIPVIA,SOLINM_NOMVIA,SOLINM_NUMVIA,SOLINM_INTDPT,SOLINM_INTDPT,SOLINM_NOMZON,SOLINM_TIPZON,SOLINM_UBIGEO,EVALEG_FEENIN"
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A"
   g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.HIPMAE_TDOCLI AND B.DATGEN_NUMDOC = A.HIPMAE_NDOCLI"
   g_str_Parame = g_str_Parame & "  INNER JOIN TRA_EVATAS C ON C.EVATAS_NUMSOL = A.HIPMAE_NUMSOL"
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 223 AND D.PARDES_CODITE = C.EVATAS_MATCON"
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 221 AND E.PARDES_CODITE = C.EVATAS_TIPINM"
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_SOLINM F ON F.SOLINM_NUMSOL = C.EVATAS_NUMSOL"
   g_str_Parame = g_str_Parame & "  INNER JOIN TRA_EVALEG G ON A.HIPMAE_NUMSOL = G.EVALEG_NUMSOL" 'añadido 11/04/2016
   g_str_Parame = g_str_Parame & "  WHERE A.HIPMAE_NUMOPE='" & moddat_g_str_NumOpe & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
      
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      For i = 1 To 19 '22
         Printer.FontSize = 10
         Printer.Print ""
      Next
      
      With g_rst_Princi
         Printer.Print Tab(30); Mid(!NOMBRE_COMPLETO, 1, 33); Tab(110); Trim(!DNI)

         'utilizamos esta lineacon tipo de fuente 4 para imprimir solo en blanco con este y luego a su tamaño normal(10)
         Printer.FontSize = 4: Printer.Print "":  Printer.FontSize = 10
         Printer.Print Tab(5); Mid(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                        " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                        IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "") & " " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo)), 1, 80)

         'de igual manera utilizamos esta linea y la imprimimos con este tamaño(4) y luego al tamaño normal(10)
         'para poder hacer ingresar Direccion como un tamaño normal que se encuentra entre ambas lineas en blanco,
         'debido a que cada vez hacia que imprima en una linea mas ancha y nos descuadraba el formato del documento
         Printer.FontSize = 4: Printer.Print "": Printer.Print "":  Printer.FontSize = 10

         Printer.Print Tab(5); Trim(!TELEFONO); Tab(65); IIf(IsNull(!CORREO), "", !CORREO)
         Printer.Print ""

         If frm_Ges_CreHip_02.grd_Listad.TextMatrix(7, 1) = "BIEN TERMINADO" Then
            Printer.Print Tab(40); Trim(!FECHA_DESEMBOLSO); Tab(90); Trim(!ULTIMA_CUOTA)
         Else
            If IsNull(!EVALEG_FEENIN) Then
               Printer.Print Tab(40); !FECHA_DESEMBOLSO; Tab(90); Trim(!ULTIMA_CUOTA)
            Else
               If (!EVALEG_FEENIN) = 0 Then
                  Printer.Print Tab(40); !FECHA_DESEMBOLSO; Tab(90); Trim(!ULTIMA_CUOTA)
               Else
                  Printer.Print Tab(40); gf_FormatoFecha(CStr(!EVALEG_FEENIN)); Tab(90); Trim(!ULTIMA_CUOTA)
               End If
            End If
         End If

         For i = 1 To 3
            Printer.Print ""
         Next

         Printer.Print Tab(20); Format(!SUMA_ASEGURADA, "#,###,##0.00")
         Printer.FontSize = 4: Printer.Print "": Printer.Print "": Printer.FontSize = 10
         Printer.Print Tab(20); Mid(moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
               " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA), 1, 35); Tab(90); moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
         Printer.Print Tab(20); moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00")
         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
         Printer.Print Tab(20); Trim(!NUMERO_PISOS); Tab(80); Trim(!NUMERO_SOTANOS)
         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10
         Printer.Print Tab(20); Trim(!ANIO_CONSTRUCCION)
         Printer.FontSize = 4: Printer.Print "":  Printer.FontSize = 10
         Printer.Print Tab(30); Trim(!CARACT_BIEN_ASEGURADO)

         For i = 1 To 7
            Printer.Print ""
         Next
         Printer.FontSize = 4: Printer.Print "": Printer.FontSize = 10

         Printer.EndDoc
      End With
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

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
   Call fs_Buscar_DatEva
   Call gs_CentraForm(Me)
   
   'Validacion para los creditos extornados y cancelados
   If moddat_g_int_Situac <> 2 Then
      cmd_Grabar.Enabled = False
   End If
   
   Call gs_SetFocus(txt_NumPol_Des)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   txt_NumPol_Des.Text = ""
   txt_NumPoC_Des.Text = ""
   txt_NumPol_Viv.Text = ""
   
   If moddat_g_int_Situac = 9 Then
      cmd_Grabar.Enabled = False
   End If
   
   'Obteniendo Información de Evaluación de Seguros
   g_str_Parame = "SELECT * FROM TRA_EVASEG WHERE "
   g_str_Parame = g_str_Parame & "EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      pnl_EmpSeg.Caption = moddat_gf_Consulta_ComSeg(g_rst_Princi!EVASEG_ESGDES & "")
      pnl_TipSeg.Caption = moddat_gf_Consulta_TipSeg(g_rst_Princi!EVASEG_ESGDES, g_rst_Princi!EVASEG_TIPSEG)
      pnl_FecEva_Des.Caption = gf_FormatoFecha(g_rst_Princi!EVASEG_EVADES)
      pnl_TipApl_Des.Caption = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPDES))
      pnl_ValApl_Des.Caption = Format(g_rst_Princi!EVASEG_FOIDES, "###,###,##0.000000")
      pnl_FecEva_Viv.Caption = gf_FormatoFecha(g_rst_Princi!EVASEG_EVAVIV)
      pnl_TipApl_Viv.Caption = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPVIV))
      pnl_ValApl_Viv.Caption = Format(g_rst_Princi!EVASEG_FOIVIV, "###,###,##0.000000")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatEva()
   moddat_g_int_FlgGrb = 1
   g_str_Parame = "SELECT * FROM TRA_POLIZA WHERE "
   g_str_Parame = g_str_Parame & "POLIZA_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_g_int_FlgGrb = 2
      g_rst_Princi.MoveFirst
      
      pnl_FecEmi_Des.Caption = gf_FormatoFecha(g_rst_Princi!POLIZA_FEMDES)
      txt_NumPol_Des.Text = Trim(g_rst_Princi!POLIZA_NUMDES & "")
      txt_NumPoC_Des.Text = Trim(g_rst_Princi!POLIZA_NUMCYG & "")
      pnl_FecEmi_Viv.Caption = gf_FormatoFecha(g_rst_Princi!POLIZA_FEMVIV)
      txt_NumPol_Viv.Text = Trim(g_rst_Princi!POLIZA_NUMVIV & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_NumPol_Des_GotFocus()
   Call gs_SelecTodo(txt_NumPol_Des)
End Sub

Private Sub txt_NumPol_Des_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumPoC_Des)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-/():;")
   End If
End Sub

Private Sub txt_NumPoC_Des_GotFocus()
   Call gs_SelecTodo(txt_NumPoC_Des)
End Sub

Private Sub txt_NumPoC_Des_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumPol_Viv)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-/():;")
   End If
End Sub

Private Sub txt_NumPol_Viv_GotFocus()
   Call gs_SelecTodo(txt_NumPol_Viv)
End Sub

Private Sub txt_NumPol_Viv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-/():;")
   End If
End Sub
