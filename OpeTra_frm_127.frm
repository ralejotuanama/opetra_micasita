VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_ModSol_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2745
   ClientLeft      =   4095
   ClientTop       =   5235
   ClientWidth     =   11595
   Icon            =   "OpeTra_frm_127.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2745
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11610
      _Version        =   65536
      _ExtentX        =   20479
      _ExtentY        =   4842
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
         Height          =   435
         Left            =   30
         TabIndex        =   4
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.ComboBox cmb_DiaPag 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   1635
         End
         Begin VB.Label Label6 
            Caption         =   "Día de Pago:"
            Height          =   315
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   6
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_127.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_127.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   7
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   720
            TabIndex        =   13
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Modificación de Solicitud de Crédito Hipotecario"
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
            Height          =   315
            Left            =   720
            TabIndex        =   14
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cambio de Día de Pago"
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
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_127.frx":0890
            Stretch         =   -1  'True
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   8
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   9
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
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
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frm_ModSol_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_DiaPag()   As moddat_tpo_Genera
Dim l_int_DiaPago    As Integer

Private Sub Grabar_Auditoria()
Dim r_str_Proceso    As String
Dim r_str_Tabla      As String
Dim r_str_Descri     As String
Dim r_str_Descri1    As String
Dim r_str_Descri2    As String
Dim r_str_Descri3    As String
Dim r_str_Usuario    As String
Dim r_str_Plataforma As String
Dim r_str_Terminal   As String
Dim r_str_Sucursal   As String

   r_str_Proceso = "CAMBIO DIA PAGO"
   r_str_Tabla = "CRE_SOLMAE"
   r_str_Usuario = modgen_g_str_CodUsu
   r_str_Terminal = modgen_g_str_NombPC
   r_str_Plataforma = UCase(App.EXEName)
   r_str_Sucursal = modgen_g_str_CodSuc
   r_str_Descri1 = ""
   r_str_Descri2 = ""
   r_str_Descri3 = ""

   'Verificacion de datos modificados para ser guardados como Auditoria
   If l_int_DiaPago <> cmb_DiaPag.Text Then
      r_str_Descri = r_str_Descri + "Dia de Pago (Antes: " & l_int_DiaPago & ")  (Nuevo: " & cmb_DiaPag.Text & ")" + Chr(13)
   End If

   r_str_Descri1 = Mid(r_str_Descri, 1, 2000)

   'Grabacion en Tabla de Auditoria
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "INSERT INTO CRE_AUDIT("
   g_str_Parame = g_str_Parame & "  AUDIT_PROCES, "
   g_str_Parame = g_str_Parame & "  AUDIT_TBLAFE, "
   g_str_Parame = g_str_Parame & "  AUDIT_NUMOPE, "
   g_str_Parame = g_str_Parame & "  AUDIT_PERIOD, "
   g_str_Parame = g_str_Parame & "  AUDIT_FECHA, "
   g_str_Parame = g_str_Parame & "  AUDIT_HORA, "
   g_str_Parame = g_str_Parame & "  AUDIT_DESCR1, "
   g_str_Parame = g_str_Parame & "  AUDIT_DESCR2, "
   g_str_Parame = g_str_Parame & "  AUDIT_DESCR3, "
   g_str_Parame = g_str_Parame & "  SEGUSUCRE, "
   g_str_Parame = g_str_Parame & "  SEGFECCRE, "
   g_str_Parame = g_str_Parame & "  SEGHORCRE, "
   g_str_Parame = g_str_Parame & "  SEGPLTCRE, "
   g_str_Parame = g_str_Parame & "  SEGTERCRE, "
   g_str_Parame = g_str_Parame & "  SEGSUCCRE) "

   g_str_Parame = g_str_Parame & "VALUES ("
   g_str_Parame = g_str_Parame & "'" & r_str_Proceso & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Tabla & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'" & Format(date, "YYYYMMDD") & "', "
   g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Descri1 & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Descri2 & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Descri3 & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Usuario & "', "
   g_str_Parame = g_str_Parame & "'" & Format(date, "YYYYMMDD") & "', "
   g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Plataforma & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Terminal & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Sucursal & "' )"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If

End Sub

Private Sub cmd_Grabar_Click()
   If cmb_DiaPag.ListIndex = -1 Then
      MsgBox "Debe ingresar el Día de Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DiaPag)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = "UPDATE CRE_SOLMAE SET "
   g_str_Parame = g_str_Parame & "SOLMAE_DIAPAG = " & CStr(CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo)) & " "
   g_str_Parame = g_str_Parame & "WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If

   Call Grabar_Auditoria
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
      
   Call moddat_gs_Carga_ParSubPrd(cmb_DiaPag, l_arr_DiaPag(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "009")
   
   'Asignacion de variables usadas en el Proceso de Auditoria.
   If cmb_DiaPag.Text = "" Then
      l_int_DiaPago = 0
   Else
      l_int_DiaPago = cmb_DiaPag.Text
   End If
   
   l_int_DiaPago = Format(modmip_g_int_DiaMor, "00") 'frm_ModSol_02.grd_Listad(3).TextMatrix(21, 1)
   cmb_DiaPag.Text = Format(modmip_g_int_DiaMor, "00") 'frm_ModSol_02.grd_Listad(3).TextMatrix(21, 1)
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_DiaPag)
   Screen.MousePointer = 0
End Sub
