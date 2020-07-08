VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Desemb_13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   4020
   ClientTop       =   4065
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_073.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3075
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11610
      _Version        =   65536
      _ExtentX        =   20479
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   30
         TabIndex        =   6
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.ComboBox cmb_BanChq 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   3225
         End
         Begin VB.TextBox txt_NumChq 
            Height          =   315
            Left            =   1590
            MaxLength       =   25
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   3225
         End
         Begin VB.ComboBox cmb_CtaChq 
            Height          =   315
            Left            =   7590
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   3225
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Index           =   11
            Left            =   5820
            TabIndex        =   9
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Banco:"
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   8
            Top             =   390
            Width           =   1365
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Cheque:"
            Height          =   285
            Index           =   16
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Picture         =   "OpeTra_frm_073.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_073.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Left            =   690
            TabIndex        =   12
            Top             =   30
            Width           =   8685
            _Version        =   65536
            _ExtentX        =   15319
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios"
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
            Height          =   315
            Left            =   690
            TabIndex        =   20
            Top             =   330
            Width           =   8685
            _Version        =   65536
            _ExtentX        =   15319
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Desembolso - Regularización de Cheque de Gerencia"
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
            Picture         =   "OpeTra_frm_073.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   13
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Left            =   1590
            TabIndex        =   14
            Top             =   60
            Width           =   6045
            _Version        =   65536
            _ExtentX        =   10663
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   9270
            TabIndex        =   15
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
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
         End
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1590
            TabIndex        =   16
            Top             =   390
            Width           =   9915
            _Version        =   65536
            _ExtentX        =   17489
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
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   7890
            TabIndex        =   18
            Top             =   90
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   17
            Top             =   90
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_Desemb_13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_BanChq()   As moddat_tpo_Genera
Dim l_arr_CtaChq()   As moddat_tpo_Genera

Private Sub cmb_BanChq_Click()
   Call gs_SetFocus(cmb_CtaChq)
   
   If cmb_BanChq.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_CtaBan(l_arr_BanChq(cmb_BanChq.ListIndex + 1).Genera_Codigo, cmb_CtaChq, l_arr_CtaChq)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_BanChq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BanChq_Click
   End If
End Sub

Private Sub cmb_CtaChq_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_CtaChq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaChq_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_NumChq.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Cheque.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumChq)
      Exit Sub
   End If
   
   If cmb_BanChq.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco del Cheque.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_BanChq)
      Exit Sub
   End If

   If cmb_CtaChq.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Cuenta del Cheque.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaChq)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'Grabando Cabecera de Credito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPDES_REGCHQ ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumChq.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_CtaChq(cmb_CtaChq.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_BanChq(cmb_BanChq.ListIndex + 1).Genera_Codigo & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                           'Código Sucursal
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop

   moddat_g_int_FlgAct = 2
   
   MsgBox "Se modificaron los datos correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_NomCli.Caption = moddat_g_str_NomCli
   
   Call fs_Inicia
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_BanChq, l_arr_BanChq, 1, "516")

   txt_NumChq.Text = ""
   cmb_BanChq.ListIndex = -1
   cmb_CtaChq.Clear
End Sub

Private Sub txt_NumChq_GotFocus()
   Call gs_SelecTodo(txt_NumChq)
End Sub

Private Sub txt_NumChq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_BanChq)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub


