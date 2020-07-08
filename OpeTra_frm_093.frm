VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_EmpPer_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3480
   ClientLeft      =   4125
   ClientTop       =   2760
   ClientWidth     =   8190
   Icon            =   "OpeTra_frm_093.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   6165
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
         Height          =   1425
         Left            =   30
         TabIndex        =   7
         Top             =   2010
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
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
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1050
            Width           =   6615
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   845
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1440
            MaxLength       =   120
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   6615
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   720
            Width           =   6615
         End
         Begin VB.Label Label8 
            Caption         =   "Situación:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "Código Contacto:"
            Height          =   285
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label4 
            Caption         =   "Nombre Contacto:"
            Height          =   285
            Left            =   60
            TabIndex        =   9
            Top             =   390
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail:"
            Height          =   285
            Left            =   60
            TabIndex        =   8
            Top             =   720
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   30
         TabIndex        =   12
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
            Picture         =   "OpeTra_frm_093.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7500
            Picture         =   "OpeTra_frm_093.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   30
         TabIndex        =   13
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
            TabIndex        =   14
            Top             =   60
            Width           =   5535
            _Version        =   65536
            _ExtentX        =   9763
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
            TabIndex        =   15
            Top             =   330
            Width           =   6315
            _Version        =   65536
            _ExtentX        =   11139
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Personas de Contacto"
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
            Picture         =   "OpeTra_frm_093.frx":0890
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   525
         Left            =   60
         TabIndex        =   16
         Top             =   1440
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   926
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
         Begin Threed.SSPanel pnl_EmpPer 
            Height          =   405
            Left            =   1440
            TabIndex        =   17
            Top             =   60
            Width           =   6585
            _Version        =   65536
            _ExtentX        =   11615
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "SSPanel3"
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
         Begin VB.Label lbl_NomEmp 
            Caption         =   "Empresa Peritaje:"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   150
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_EmpPer_06"
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
      txt_Codigo.Text = Format(txt_Codigo.Text, "000")
      
      If Len(Trim(txt_Codigo.Text)) = 0 Then
         MsgBox "Debe ingresar el Código del Contacto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
         Exit Sub
      End If
      
      If txt_Codigo.Text = "000" Then
         MsgBox "El Código ingresado es incorrecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre del Contacto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If

   If Len(Trim(txt_DirEle.Text)) = 0 Then
      MsgBox "Debe ingresar el E-Mail.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DirEle)
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
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM MNT_PERCON "
      g_str_Parame = g_str_Parame & " WHERE PERCON_CODEMP = '" & moddat_g_str_CodGrp & "'"
      g_str_Parame = g_str_Parame & "   AND PERCON_CODCON = '" & txt_Codigo.Text & "' "
      
      'Tipo de empresa
      If moddat_g_int_TipRec = 1 Then
         g_str_Parame = g_str_Parame & " AND PERCON_TIPTAB = 1 "
      ElseIf moddat_g_int_TipRec = 2 Then
         g_str_Parame = g_str_Parame & " AND PERCON_TIPTAB = 2 "
      ElseIf moddat_g_int_TipRec = 3 Then
         g_str_Parame = g_str_Parame & " AND PERCON_TIPTAB = 3 "
      End If
   
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
      
      g_str_Parame = "USP_MNT_PERCON ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_DirEle.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
      
      'Tipo de empresa
      If moddat_g_int_TipRec = 1 Then
         g_str_Parame = g_str_Parame & "1, "
      ElseIf moddat_g_int_TipRec = 2 Then
         g_str_Parame = g_str_Parame & "2, "
      ElseIf moddat_g_int_TipRec = 3 Then
         g_str_Parame = g_str_Parame & "3, "
      End If
   
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
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   moddat_g_int_FlgAct_1 = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Call gs_CentraForm(Me)
   Me.Caption = modgen_g_str_NomPlt
   
   'Titulos
   If moddat_g_int_TipRec = 1 Then
      pnl_TitPri.Caption = "Empresas de Peritaje"
      lbl_NomEmp.Caption = "Empresa Peritaje:"
   ElseIf moddat_g_int_TipRec = 2 Then
      pnl_TitPri.Caption = "Empresas de Seguros"
      lbl_NomEmp.Caption = "Empresa Seguros:"
   ElseIf moddat_g_int_TipRec = 3 Then
      pnl_TitPri.Caption = "Empresas de Notaria"
      lbl_NomEmp.Caption = "Notaria:"
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      pnl_TitSec.Caption = "Personas de Contacto - Nuevo Registro"
   Else
      pnl_TitSec.Caption = "Personas de Contacto - Modificación de Datos"
   End If
   
   pnl_EmpPer.Caption = moddat_g_str_CodGrp & " - " & moddat_g_str_DesGrp

   'Limpiando Variables
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "244")
   txt_Codigo.Text = ""
   txt_Nombre.Text = ""
   txt_DirEle.Text = ""
   
   Call gs_SetFocus(txt_Codigo)

   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT * "
      g_str_Parame = g_str_Parame & "   FROM MNT_PERCON "
      g_str_Parame = g_str_Parame & "  WHERE PERCON_CODEMP = '" & moddat_g_str_CodGrp & "'"
      g_str_Parame = g_str_Parame & "    AND PERCON_CODCON = '" & moddat_g_str_Codigo & "' "
      
      'Tipo de empresa
      If moddat_g_int_TipRec = 1 Then
         g_str_Parame = g_str_Parame & " AND PERCON_TIPTAB = 1 "
      ElseIf moddat_g_int_TipRec = 2 Then
         g_str_Parame = g_str_Parame & " AND PERCON_TIPTAB = 2 "
      ElseIf moddat_g_int_TipRec = 3 Then
         g_str_Parame = g_str_Parame & " AND PERCON_TIPTAB = 3 "
      End If

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If

      g_rst_Genera.MoveFirst
   
      txt_Codigo.Text = Trim(g_rst_Genera!PERCON_CODCON)
      txt_Nombre.Text = Trim(g_rst_Genera!PERCON_NOMBRE)
      txt_DirEle.Text = Trim(g_rst_Genera!PERCON_DIRELE)
   
      Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Genera!PERCON_SITUAC)
      
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
      Call gs_SetFocus(txt_DirEle)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "()-_ .,;:¿?/&%$@#")
   End If
End Sub

Private Sub txt_DirEle_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Situac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_@.")
   End If
End Sub

