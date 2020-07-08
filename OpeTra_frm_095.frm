VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_EvaTas_13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4305
   ClientLeft      =   1755
   ClientTop       =   3555
   ClientWidth     =   11250
   Icon            =   "OpeTra_frm_095.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4290
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   7567
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
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   3810
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
         Begin VB.TextBox txt_NomArc 
            Height          =   315
            Left            =   1860
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   60
            Width           =   8745
         End
         Begin VB.CommandButton cmd_BusArc 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10680
            TabIndex        =   21
            Top             =   60
            Width           =   435
         End
         Begin VB.Label Label4 
            Caption         =   "Archivo a adjuntar:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   435
         Left            =   30
         TabIndex        =   1
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
         Begin VB.TextBox txt_Asunto 
            Height          =   315
            Left            =   1860
            MaxLength       =   250
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   60
            Width           =   9255
         End
         Begin VB.Label Label7 
            Caption         =   "Asunto:"
            Height          =   285
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1035
         Left            =   30
         TabIndex        =   4
         Top             =   2730
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
         Begin VB.TextBox txt_Conten 
            Height          =   915
            Left            =   1860
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Text            =   "OpeTra_frm_095.frx":000C
            Top             =   60
            Width           =   9255
         End
         Begin VB.Label Label6 
            Caption         =   "Mensaje:"
            Height          =   285
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   7
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
            TabIndex        =   8
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
               Size            =   9.76
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
            TabIndex        =   9
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
               Size            =   9.76
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
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   8370
            Top             =   30
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10020
            Top             =   60
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
            Left            =   9390
            Top             =   60
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
            Picture         =   "OpeTra_frm_095.frx":0010
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   16
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8040
            TabIndex        =   15
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   17
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "OpeTra_frm_095.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_EnvCor 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_095.frx":075C
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Enviar por Correo Electrónico"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_EvaTas_13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_EmpPer     As String
Dim l_str_CorCon     As String
Dim l_str_NomCon     As String

Private Sub cmd_BusArc_Click()
   On Error GoTo cmd_BusArc_Error

   dlg_Guarda.Filter = "Archivos PDF (*.pdf)|*.pdf"
   dlg_Guarda.ShowOpen

   txt_NomArc.Text = UCase(dlg_Guarda.FileName)
   
   Exit Sub
   
cmd_BusArc_Error:
   txt_NomArc.Text = ""
   
End Sub

Private Sub cmd_EnvCor_Click()
   If Len(Trim(txt_Asunto.Text)) = 0 Then
      MsgBox "Debe ingresar el Asunto del Correo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Asunto)
      Exit Sub
   End If

   If Len(Trim(txt_Conten.Text)) = 0 Then
      MsgBox "Debe ingresar el Mensaje.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Conten)
      Exit Sub
   End If
   
   If Len(Trim(txt_NomArc.Text)) = 0 Then
      MsgBox "Debe ingresar el Archivo a adjuntar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If
   
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
   Call fs_CorEle
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_CorEle()
   txt_Asunto.Text = "ORDEN DE TASACION: " & gf_Formato_NumSol(moddat_g_str_NumSol) & " / " & CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   txt_Conten.Text = ""
   txt_Conten.Text = txt_Conten.Text & "SEÑORES: " & l_str_EmpPer & Chr(10)
   txt_Conten.Text = txt_Conten.Text & Chr(10)
   txt_Conten.Text = txt_Conten.Text & "         " & l_str_NomCon & Chr(10) & Chr(10)
   txt_Conten.Text = txt_Conten.Text & "POR MEDIO DEL PRESENTE ADJUNTO ORDEN DE TASACION DEL CLIENTE DE LA REFERENCIA." & Chr(13)
End Sub

Private Sub fs_Inicia()
   txt_Asunto.Text = ""
   txt_Conten.Text = ""
   txt_NomArc.Text = ""

   l_str_EmpPer = ""
   l_str_CorCon = ""
   l_str_NomCon = ""

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM RPT_ORDTAS WHERE "
   g_str_Parame = g_str_Parame & "ORDTAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      
      l_str_EmpPer = moddat_gf_Consulta_ParDes("507", g_rst_Genera!ORDTAS_EMPPER)
      l_str_NomCon = moddat_gf_Consulta_PerCon(g_rst_Genera!ORDTAS_EMPPER, g_rst_Genera!ORDTAS_PERCON, l_str_CorCon)
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub txt_Asunto_GotFocus()
   Call gs_SelecTodo(txt_Asunto)
End Sub

Private Sub txt_Asunto_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .()/#$@")
   Else
      Call gs_SetFocus(txt_Conten)
   End If
End Sub

Private Sub txt_Conten_GotFocus()
   Call gs_SelecTodo(txt_Conten)
End Sub

Private Sub txt_Conten_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomArc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_NomArc_GotFocus()
   Call gs_SelecTodo(txt_NomArc)
End Sub

Private Sub txt_NomArc_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_.\")
   Else
      Call gs_SetFocus(cmd_EnvCor)
   End If
End Sub
