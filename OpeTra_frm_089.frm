VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_EmpPer_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2580
   ClientLeft      =   3450
   ClientTop       =   4260
   ClientWidth     =   8175
   Icon            =   "OpeTra_frm_089.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2595
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   4577
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
         Height          =   1095
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
         _ExtentY        =   1931
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
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   6615
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   845
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1440
            MaxLength       =   80
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   6615
         End
         Begin VB.Label Label8 
            Caption         =   "Situación:"
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   750
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "Código Empresa:"
            Height          =   285
            Left            =   90
            TabIndex        =   11
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label Label4 
            Caption         =   "Razón Social:"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   420
            Width           =   1305
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   30
         TabIndex        =   7
         Top             =   750
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
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
            Picture         =   "OpeTra_frm_089.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7500
            Picture         =   "OpeTra_frm_089.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   270
            Left            =   630
            TabIndex        =   9
            Top             =   60
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Empresas de Peritaje"
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
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   270
            Left            =   630
            TabIndex        =   13
            Top             =   330
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Nuevo Registro"
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
            Left            =   90
            Picture         =   "OpeTra_frm_089.frx":0890
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_EmpPer_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If moddat_g_int_FlgGrb = 1 Then
      txt_Codigo.Text = Format(txt_Codigo.Text, "000000")
      
      If Len(Trim(txt_Codigo.Text)) = 0 Then
         MsgBox "Debe ingresar el Código de la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
         Exit Sub
      End If
      
      If txt_Codigo.Text = "000000" Then
         MsgBox "El Código ingresado es incorrecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
         Exit Sub
      End If
   End If
   
   
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If

   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If

   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP = '507' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODITE = '" & txt_Codigo.Text & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
        g_rst_Genera.Close
        Set g_rst_Genera = Nothing
        
        MsgBox "El Código ya ha sido registrado. Por favor verifique el código e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(txt_Codigo)
        Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      If moddat_g_int_FlgGrb = 1 Then
         g_str_Parame = "USP_INSERTA_MNT_PARDES ("
         g_str_Parame = g_str_Parame & "'507', "
         g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
         g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      Else
         g_str_Parame = "USP_MODIFICA_MNT_PARDES ("
         g_str_Parame = g_str_Parame & "'507', "
         g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
         g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      End If
      
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
      
      Screen.MousePointer = 0
   Loop
   
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Call gs_CentraForm(Me)
   Me.Caption = modgen_g_str_NomPlt
   
   If moddat_g_int_TipRec = 1 Then
      pnl_TitPri.Caption = "Empresas de Peritaje"
   ElseIf moddat_g_int_TipRec = 2 Then
      pnl_TitPri.Caption = "Empresas de Seguros"
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      pnl_TitSec.Caption = "Nuevo Registro"
   Else
      pnl_TitSec.Caption = "Modificación de Datos"
   End If

   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "244")
   txt_Codigo.Text = ""
   txt_Nombre.Text = ""
   
   Call gs_SetFocus(txt_Codigo)

   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP = '507' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODITE = '" & moddat_g_str_Codigo & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If

      g_rst_Genera.MoveFirst
   
      txt_Codigo.Text = Trim(g_rst_Genera!PARDES_CODITE)
      txt_Nombre.Text = Trim(g_rst_Genera!PARDES_DESCRI)
   
      Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Genera!PARDES_SITUAC)
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      txt_Codigo.Enabled = False
      
      Call gs_SetFocus(txt_Nombre)
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub txt_Codigo_GotFocus()
   Call gs_SelecTodo(txt_Codigo)
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Situac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "()-_ .,;:¿?/&%$@#")
   End If
End Sub
