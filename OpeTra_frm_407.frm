VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_EmpPer_09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11355
   Icon            =   "OpeTra_frm_407.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3675
      Left            =   -30
      TabIndex        =   7
      Top             =   0
      Width           =   11445
      _Version        =   65536
      _ExtentX        =   20188
      _ExtentY        =   6482
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   525
         Left            =   45
         TabIndex        =   8
         Top             =   1470
         Width           =   11265
         _Version        =   65536
         _ExtentX        =   19870
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
            Left            =   1260
            TabIndex        =   9
            Top             =   60
            Width           =   9915
            _Version        =   65536
            _ExtentX        =   17489
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
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Left            =   330
            TabIndex        =   10
            Top             =   150
            Width           =   660
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   45
         TabIndex        =   11
         Top             =   780
         Width           =   11265
         _Version        =   65536
         _ExtentX        =   19870
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
            Left            =   10650
            Picture         =   "OpeTra_frm_407.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_407.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Listado"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   45
         TabIndex        =   12
         Top             =   60
         Width           =   11265
         _Version        =   65536
         _ExtentX        =   19870
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
            TabIndex        =   13
            Top             =   180
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Datos de la Empresas de Seguros"
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
            Left            =   10800
            Top             =   30
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
            Left            =   90
            Picture         =   "OpeTra_frm_407.frx":0890
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1515
         Left            =   45
         TabIndex        =   14
         Top             =   2040
         Width           =   11265
         _Version        =   65536
         _ExtentX        =   19870
         _ExtentY        =   2672
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
         Begin VB.TextBox txt_DirEle2 
            Height          =   315
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   270
            Width           =   4140
         End
         Begin VB.TextBox txt_DirEle1 
            Height          =   315
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   270
            Width           =   4140
         End
         Begin VB.TextBox txt_DirEle3 
            Height          =   315
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   600
            Width           =   4140
         End
         Begin VB.TextBox txt_DirEle4 
            Height          =   315
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   600
            Width           =   4140
         End
         Begin VB.TextBox txt_DirEle5 
            Height          =   315
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   930
            Width           =   4140
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Correo 2:"
            Height          =   195
            Left            =   5820
            TabIndex        =   19
            Top             =   330
            Width           =   645
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Correo 1:"
            Height          =   195
            Left            =   330
            TabIndex        =   18
            Top             =   330
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Correo 3:"
            Height          =   195
            Left            =   330
            TabIndex        =   17
            Top             =   660
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Correo 4:"
            Height          =   195
            Left            =   5820
            TabIndex        =   16
            Top             =   660
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Correo 5:"
            Height          =   195
            Left            =   330
            TabIndex        =   15
            Top             =   990
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frm_EmpPer_09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_Codigo  As String

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call gs_CentraForm(Me)
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   pnl_EmpPer.Caption = moddat_g_str_CodGrp & " - " & moddat_g_str_DesGrp
   txt_DirEle1.Text = ""
   txt_DirEle2.Text = ""
   txt_DirEle3.Text = ""
   txt_DirEle4.Text = ""
   txt_DirEle5.Text = ""
End Sub

Private Sub fs_Buscar()
Dim r_str_Parame    As String
Dim r_rst_Princi    As ADODB.Recordset
   
   l_str_Codigo = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT DATEMP_CODEMP, DATEMP_DIRELE1, DATEMP_DIRELE2, "
   r_str_Parame = r_str_Parame & "        DATEMP_DIRELE3, DATEMP_DIRELE4, DATEMP_DIRELE5 "
   r_str_Parame = r_str_Parame & "   FROM MNT_DATEMP "
   r_str_Parame = r_str_Parame & "  WHERE DATEMP_CODEMP = '" & moddat_g_str_CodGrp & "' "
   r_str_Parame = r_str_Parame & "    AND DATEMP_TIPTAB = 3 "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
         
      l_str_Codigo = r_rst_Princi!DATEMP_CODEMP
      txt_DirEle1.Text = Trim(r_rst_Princi!DATEMP_DIRELE1 & "")
      txt_DirEle2.Text = Trim(r_rst_Princi!DATEMP_DIRELE2 & "")
      txt_DirEle3.Text = Trim(r_rst_Princi!DATEMP_DIRELE3 & "")
      txt_DirEle4.Text = Trim(r_rst_Princi!DATEMP_DIRELE4 & "")
      txt_DirEle5.Text = Trim(r_rst_Princi!DATEMP_DIRELE5 & "")
   End If
      
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_Parame     As String
Dim r_rst_Genera     As ADODB.Recordset
Dim r_int_Fila       As Integer

   If Len(Trim(moddat_g_str_CodGrp)) = 0 Then
      MsgBox "Debe de seleccionar una empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DirEle1)
      Exit Sub
   End If

   If Len(Trim(pnl_EmpPer.Caption)) = 0 Then
      MsgBox "Debe de seleccionar una empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DirEle1)
      Exit Sub
   End If
   
   If Len(Trim(txt_DirEle1.Text)) = 0 Then
      MsgBox "Debe se ingresar un correo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DirEle1)
      Exit Sub
   Else
      If gf_ValidarEmail(txt_DirEle1) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle1)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_DirEle2.Text)) > 0 Then
      If gf_ValidarEmail(txt_DirEle2) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle2)
         Exit Sub
      End If
   End If
'----------------------------
   If Len(Trim(txt_DirEle3.Text)) > 0 Then
      If gf_ValidarEmail(txt_DirEle3) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle3)
         Exit Sub
      End If
   End If
   If Len(Trim(txt_DirEle4.Text)) > 0 Then
      If gf_ValidarEmail(txt_DirEle4) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle4)
         Exit Sub
      End If
   End If
   If Len(Trim(txt_DirEle5.Text)) > 0 Then
      If gf_ValidarEmail(txt_DirEle5) = False Then
         MsgBox "El correo ingresado es erroneo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_DirEle5)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "USP_MNT_DATEMP ("
   r_str_Parame = r_str_Parame & "'" & moddat_g_str_CodGrp & "',"
   r_str_Parame = r_str_Parame & "3," 'Tipo Tabla
   r_str_Parame = r_str_Parame & "0,"
   r_str_Parame = r_str_Parame & "'" & Trim(txt_DirEle1.Text) & "',"
   r_str_Parame = r_str_Parame & "'" & Trim(txt_DirEle2.Text) & "',"
   r_str_Parame = r_str_Parame & "'" & Trim(txt_DirEle3.Text) & "',"
   r_str_Parame = r_str_Parame & "'" & Trim(txt_DirEle4.Text) & "',"
   r_str_Parame = r_str_Parame & "'" & Trim(txt_DirEle5.Text) & "',"
   r_str_Parame = r_str_Parame & "1,"
   'Datos de Auditoria
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
   r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   r_str_Parame = r_str_Parame & "1) "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 2) Then
      Exit Sub
   End If
  
   MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub txt_DirEle1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub

Private Sub txt_DirEle2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle3)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub

Private Sub txt_DirEle3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle4)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub

Private Sub txt_DirEle4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle5)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub

Private Sub txt_DirEle5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "@_.-")
   End If
End Sub
