VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Tra_EvaSeg_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   Icon            =   "OpeTra_frm_353.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8190
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11010
      _Version        =   65536
      _ExtentX        =   19420
      _ExtentY        =   14446
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   5865
         Left            =   30
         TabIndex        =   21
         Top             =   2250
         Width           =   10845
         _Version        =   65536
         _ExtentX        =   19129
         _ExtentY        =   10345
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
         Begin VB.TextBox txt_Mensaje 
            Height          =   3050
            Left            =   1470
            MaxLength       =   8000
            MultiLine       =   -1  'True
            TabIndex        =   37
            Text            =   "OpeTra_frm_353.frx":000C
            Top             =   2730
            Width           =   9255
         End
         Begin VB.TextBox txt_Asunto 
            Height          =   315
            Left            =   1470
            MaxLength       =   400
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   1290
            Width           =   9255
         End
         Begin VB.TextBox txt_Destino 
            Height          =   640
            Left            =   1470
            MaxLength       =   600
            MultiLine       =   -1  'True
            TabIndex        =   1
            Text            =   "OpeTra_frm_353.frx":0012
            Top             =   600
            Width           =   9255
         End
         Begin VB.TextBox txt_Adjunto1 
            Height          =   315
            Left            =   1470
            MaxLength       =   150
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1650
            Width           =   4850
         End
         Begin VB.TextBox txt_Adjunto2 
            Height          =   315
            Left            =   1470
            MaxLength       =   150
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   2010
            Width           =   4850
         End
         Begin VB.CommandButton cmd_Adjun1 
            Caption         =   "...."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9600
            TabIndex        =   5
            Top             =   1650
            Width           =   525
         End
         Begin VB.CommandButton cmd_Previa1 
            Caption         =   "Ver"
            Height          =   315
            Left            =   10200
            TabIndex        =   6
            Top             =   1650
            Width           =   525
         End
         Begin VB.CommandButton cmd_Adjun2 
            Caption         =   "...."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9600
            TabIndex        =   9
            Top             =   2010
            Width           =   525
         End
         Begin VB.CommandButton cmd_Previa2 
            Caption         =   "Ver"
            Height          =   315
            Left            =   10200
            TabIndex        =   10
            Top             =   2010
            Width           =   525
         End
         Begin VB.CommandButton cmd_Previa3 
            Caption         =   "Ver"
            Height          =   315
            Left            =   10200
            TabIndex        =   14
            Top             =   2370
            Width           =   525
         End
         Begin VB.CommandButton cmd_Adjun3 
            Caption         =   "...."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9600
            TabIndex        =   13
            Top             =   2370
            Width           =   525
         End
         Begin VB.TextBox txt_Adjunto3 
            Height          =   315
            Left            =   1470
            MaxLength       =   150
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   2370
            Width           =   4850
         End
         Begin VB.ComboBox cmb_EmpSeg 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   9255
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   6360
            TabIndex        =   4
            Top             =   1650
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   315
            Left            =   6360
            TabIndex        =   8
            Top             =   2010
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   315
            Left            =   6360
            TabIndex        =   12
            Top             =   2370
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mensaje:"
            Height          =   195
            Left            =   210
            TabIndex        =   36
            Top             =   3915
            Width           =   645
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Adjunto 1:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1710
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Adjunto 2:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   2070
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Adjunto 3:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   2430
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Destino:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   870
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Asunto:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1350
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Empresa Seguros:"
            Height          =   195
            Left            =   60
            TabIndex        =   22
            Top             =   300
            Width           =   1290
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   23
         Top             =   30
         Width           =   10850
         _Version        =   65536
         _ExtentX        =   19138
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
            TabIndex        =   24
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Evaluación de Seguros"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   630
            TabIndex        =   25
            Top             =   330
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Envió de Correo"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10230
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
            Left            =   9660
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   9180
            Top             =   30
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_353.frx":0018
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   26
         Top             =   1440
         Width           =   10850
         _Version        =   65536
         _ExtentX        =   19138
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
            Left            =   1470
            TabIndex        =   17
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
         Begin Threed.SSPanel pnl_FecSol 
            Height          =   315
            Left            =   9300
            TabIndex        =   19
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
            Left            =   1470
            TabIndex        =   18
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   60
            TabIndex        =   29
            Top             =   450
            Width           =   525
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Solicitud:"
            Height          =   195
            Left            =   7980
            TabIndex        =   28
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Solicitud"
            Height          =   195
            Left            =   60
            TabIndex        =   27
            Top             =   120
            Width           =   945
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   30
         Top             =   750
         Width           =   10850
         _Version        =   65536
         _ExtentX        =   19138
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
            Left            =   10230
            Picture         =   "OpeTra_frm_353.frx":0322
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_353.frx":0764
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_EvaSeg_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type arr_DatEmp
   regEmp_CodEmp      As String
   regEmp_Direle1     As String
   regEmp_Direle2     As String
   regEmp_Direle3     As String
   regEmp_Direle4     As String
   regEmp_Direle5     As String
End Type
Dim l_arr_DatEmp()      As arr_DatEmp

Private Sub cmb_EmpSeg_Click()
Dim l_int_Fila    As Integer
Dim r_str_CadAux  As String
   'Destinatarios de Correo
   txt_Destino.Text = ""
   r_str_CadAux = ""
   ReDim moddat_g_arr_Genera(0)
   
   If cmb_EmpSeg.ListIndex > -1 Then
      For l_int_Fila = 1 To UBound(l_arr_DatEmp)
          If cmb_EmpSeg.ItemData(cmb_EmpSeg.ListIndex) = l_arr_DatEmp(l_int_Fila).regEmp_CodEmp Then
             If Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle1) <> "" Then
                r_str_CadAux = Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle1) & "; "
                ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
                moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle1)
             End If
               
             If Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle2) <> "" Then
                r_str_CadAux = r_str_CadAux & Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle2) & "; "
                ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
                moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle2)
             End If
               
             If Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle3) <> "" Then
                r_str_CadAux = r_str_CadAux & Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle3) & "; "
                ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
                moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle3)
             End If
               
             If Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle4) <> "" Then
                r_str_CadAux = r_str_CadAux & Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle4) & "; "
                ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
                moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle4)
             End If
               
             If Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle5) <> "" Then
                r_str_CadAux = r_str_CadAux & Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle5) & "; "
                ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
                moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim(l_arr_DatEmp(l_int_Fila).regEmp_Direle5)
             End If
             txt_Destino.Text = r_str_CadAux
             Exit For
          End If
      Next
   End If
End Sub

Private Sub cmb_EmpSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Asunto)
   End If
End Sub

Private Sub cmd_Adjun1_Click()
  On Error GoTo cmd_BusArc_Error
 
   dlg_Guarda.Filter = "Todos los archivo PDF (*.*)|*.pdf"
   dlg_Guarda.ShowOpen
   txt_Adjunto1.Tag = dlg_Guarda.FileName
   txt_Adjunto1.Text = dlg_Guarda.FileTitle
   Exit Sub

cmd_BusArc_Error:
   txt_Adjunto1.Text = ""
   txt_Adjunto1.Tag = ""
End Sub

Private Sub cmd_Adjun2_Click()
  On Error GoTo cmd_BusArc_Error
 
   dlg_Guarda.Filter = "Todos los archivo PDF (*.*)|*.pdf"
   dlg_Guarda.ShowOpen
   txt_Adjunto2.Tag = dlg_Guarda.FileName
   txt_Adjunto2.Text = dlg_Guarda.FileTitle
   Exit Sub

cmd_BusArc_Error:
   txt_Adjunto2.Text = ""
   txt_Adjunto2.Tag = ""
End Sub

Private Sub cmd_Adjun3_Click()
  On Error GoTo cmd_BusArc_Error
 
   dlg_Guarda.Filter = "Todos los archivo PDF (*.*)|*.pdf"
   dlg_Guarda.ShowOpen
   txt_Adjunto3.Tag = dlg_Guarda.FileName
   txt_Adjunto3.Text = dlg_Guarda.FileTitle
   Exit Sub

cmd_BusArc_Error:
   txt_Adjunto3.Text = ""
   txt_Adjunto3.Tag = ""
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_Parame    As String
Dim r_rst_Princi    As ADODB.Recordset
Dim r_str_NumFil    As Integer
Dim r_int_NumLOG    As Integer
Dim r_str_ResExt    As String
Dim r_str_NomAch    As String
Dim r_str_Mensaj    As String
      
   If cmb_EmpSeg.ListIndex = -1 Then
      MsgBox "Debe de seleccionar una empresa de seguros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EmpSeg)
      Exit Sub
   End If
   
   If Trim(txt_Destino.Text) = "" Then
      MsgBox "Debe ingresar los correos de destino en el mantenedor de seguros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Asunto)
      Exit Sub
   End If
   
   If Trim(txt_Asunto.Text) = "" Then
      MsgBox "Debe ingresar un asunto para el envio de correo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Asunto)
      Exit Sub
   End If
      
   If Trim(txt_Mensaje.Text) = "" Then
      MsgBox "Debe de ingresar un mensaje para el envió de correo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Asunto)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de enviar el correo?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   r_int_NumLOG = FreeFile
   
   Open g_str_RutLogSeg & "\" & moddat_g_str_NumSol & "_42_0_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".LOG" For Output As r_int_NumLOG
   Print #r_int_NumLOG, "Proceso           : " & modgen_g_str_NomPlt
   Print #r_int_NumLOG, "Nombre Ejecutable : " & UCase(App.EXEName)
   Print #r_int_NumLOG, "Número Revisión   : " & modgen_g_str_NumRev
   Print #r_int_NumLOG, "Nombre PC         : " & modgen_g_str_NombPC
   Print #r_int_NumLOG, "Usuario Sistema   : " & modgen_g_str_CodUsu
   Print #r_int_NumLOG, "Origen Datos      : " & moddat_g_str_NomEsq & " - " & moddat_g_str_EntDat
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, "Inicio Proceso    : " & Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss")
   Print #r_int_NumLOG, "Fecha Proceso     : " & Format(date, "dd/mm/yyyy")
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, "*************************************************************************************"
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, "       Eventos"
   Print #r_int_NumLOG, "       ======="
   Print #r_int_NumLOG, ""
'----------------------------------------------------------------------
    
   'Renombrando adjunto 1 y transfiriendolo a la carpeta de seguros
   r_str_ResExt = ""
   r_str_ResExt = fs_Extension(txt_Adjunto1.Text)
   If Len(Trim(txt_Adjunto1.Tag)) > 0 Then
      r_str_NomAch = moddat_g_str_NumSol & "_42_1_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & "." & r_str_ResExt
      
      FileCopy txt_Adjunto1.Tag, g_str_RutLogSeg & r_str_NomAch
      Print #r_int_NumLOG, "*:     Transfiriendo el Adjunto_1(" & txt_Adjunto1.Tag & ") a la carpeta de tasacion, renombrandolo."
      Print #r_int_NumLOG, "       Nuevo nombre del archivo: " & g_str_RutLogSeg & r_str_NomAch
   Else
      Print #r_int_NumLOG, "*:     No hay ningún archivo a adjuntar en el correo en Adjuntar_1"
   End If
   
   'Renombrando adjunto 2 y transfiriendolo a la carpeta seguros
   r_str_ResExt = ""
   r_str_ResExt = fs_Extension(txt_Adjunto2.Text)
   If Len(Trim(txt_Adjunto2.Tag)) > 0 Then
      r_str_NomAch = moddat_g_str_NumSol & "_42_2_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & "." & r_str_ResExt
      
      FileCopy txt_Adjunto2.Tag, g_str_RutLogSeg & r_str_NomAch
      Print #r_int_NumLOG, "*:     Transfiriendo el Adjunto_2(" & txt_Adjunto2.Tag & ") a la carpeta de tasacion, renombrandolo."
      Print #r_int_NumLOG, "       Nuevo nombre del archivo: " & g_str_RutLogSeg & r_str_NomAch
   Else
      Print #r_int_NumLOG, "*:     No hay ningún archivo a adjuntar en el correo en Adjuntar_2"
   End If
    
   'Renombrando adjunto 3 y transfiriendolo a la carpeta de seguros
   r_str_ResExt = ""
   r_str_ResExt = fs_Extension(txt_Adjunto3.Text)
   If Len(Trim(txt_Adjunto3.Tag)) > 0 Then
      r_str_NomAch = moddat_g_str_NumSol & "_42_3_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & "." & r_str_ResExt
      
      FileCopy txt_Adjunto3.Tag, g_str_RutLogSeg & r_str_NomAch
      Print #r_int_NumLOG, "*:     Transfiriendo el Adjunto_3(" & txt_Adjunto3.Tag & ") a la carpeta de tasacion, renombrandolo."
      Print #r_int_NumLOG, "       Nuevo nombre del archivo: " & g_str_RutLogSeg & r_str_NomAch
   Else
      Print #r_int_NumLOG, "*:     No hay ningún archivo a adjuntar en el correo en Adjuntar_2"
   End If
   
   'Adjuntar Correos
   ReDim moddat_g_arr_GenAux(0)
   If Len(Trim(txt_Adjunto1.Text)) > 0 Then
      ReDim Preserve moddat_g_arr_GenAux(UBound(moddat_g_arr_GenAux) + 1)
      moddat_g_arr_GenAux(UBound(moddat_g_arr_GenAux)).Genera_Codigo = Trim(txt_Adjunto1.Text)
      moddat_g_arr_GenAux(UBound(moddat_g_arr_GenAux)).Genera_Refere = Trim(txt_Adjunto1.Tag)
   End If
   If Len(Trim(txt_Adjunto2.Text)) > 0 Then
      ReDim Preserve moddat_g_arr_GenAux(UBound(moddat_g_arr_GenAux) + 1)
      moddat_g_arr_GenAux(UBound(moddat_g_arr_GenAux)).Genera_Codigo = Trim(txt_Adjunto2.Text)
      moddat_g_arr_GenAux(UBound(moddat_g_arr_GenAux)).Genera_Refere = Trim(txt_Adjunto2.Tag)
   End If
   If Len(Trim(txt_Adjunto3.Text)) > 0 Then
      ReDim Preserve moddat_g_arr_GenAux(UBound(moddat_g_arr_GenAux) + 1)
      moddat_g_arr_GenAux(UBound(moddat_g_arr_GenAux)).Genera_Codigo = Trim(txt_Adjunto3.Text)
      moddat_g_arr_GenAux(UBound(moddat_g_arr_GenAux)).Genera_Refere = Trim(txt_Adjunto3.Tag)
   End If
   
   'Envio de correo
   r_str_Mensaj = ""
   r_str_Mensaj = fs_EnviarCorreoAdj(mps_Sesion, mps_Mensaj, txt_Asunto.Text, txt_Mensaje.Text, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, False, False, False)
   
   If r_str_Mensaj = "" Then
      Print #r_int_NumLOG, "*:     Se envio el correo correctamente."
   Else
      Print #r_int_NumLOG, "*:     " & r_str_Mensaj
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If r_str_Mensaj = "" Then
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 42, 94, 0, "", 0, 0) Then
         Print #r_int_NumLOG, "*:     Error al insertar la ocurrencia en el seguimiento."
         Exit Sub
      Else
         Print #r_int_NumLOG, "*:     Se inserto la ocurrencia correctamente al seguimiento."
      End If
   End If
   
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, "       Envio Correo de Evaluacion Seguros "
   Print #r_int_NumLOG, "       ================================== "
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, "*:     Nro Solicitud  : " & Trim(pnl_NumSol.Caption)
   Print #r_int_NumLOG, "*:     Cliente        : " & Trim(pnl_Client.Caption)
   Print #r_int_NumLOG, "*:     Destino        : " & Trim(txt_Destino.Text)
   Print #r_int_NumLOG, "*:     Asunto         : " & Trim(txt_Asunto.Text)
   Print #r_int_NumLOG, "*:     Adjunto 1      : " & Trim(txt_Adjunto1.Tag)
   Print #r_int_NumLOG, "*:     Adjunto 2      : " & Trim(txt_Adjunto2.Tag)
   Print #r_int_NumLOG, "*:     Adjunto 3      : " & Trim(txt_Adjunto3.Tag)
   Print #r_int_NumLOG, "*:     Mensaje        : " & Trim(txt_Mensaje.Text)
   
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, "Fin Proceso       : " & Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss")
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, ""
      
   Close #r_int_NumLOG
   If r_str_Mensaj = "" Then
      MsgBox "Se envió el correo de seguros.", vbInformation, modgen_g_str_NomPlt
   End If
   
   Screen.MousePointer = 0
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Function fs_Extension(ByVal p_NomArchivo As String) As String
Dim r_int_PosExt As Integer
Dim r_str_ResExt As String

   fs_Extension = ""
   r_int_PosExt = 0
   r_str_ResExt = ""
   If Len(Trim(p_NomArchivo)) > 0 Then
       r_int_PosExt = InStrRev(p_NomArchivo, ".")
       If r_int_PosExt <> 0 Then
         r_str_ResExt = Right(p_NomArchivo, Len(p_NomArchivo) - r_int_PosExt)
       Else
         r_str_ResExt = ""
       End If
    End If
        
    fs_Extension = r_str_ResExt
End Function

'Private Function fs_EnviarCorreo(p_Sesion As MAPISession, p_Mensaje As MAPIMessages, p_Arregl() As moddat_tpo_Genera, p_Asunto As String, p_Contenido As String) As String
'
'   Dim r_int_Contad      As Integer
'   Dim r_int_Index       As Integer
'
'   On Error GoTo moddat_gf_EnvCor
'
'   fs_EnviarCorreo = ""
'   'Inicializa
'   p_Sesion.DownLoadMail = False
'   p_Sesion.NewSession = True
'   p_Sesion.SignOn
'   p_Mensaje.SessionID = p_Sesion.SessionID
'
'   'Envío
'   p_Mensaje.Compose
'
'   For r_int_Contad = 0 To UBound(p_Arregl) - 1
'      If Len(Trim(p_Arregl(r_int_Contad + 1).Genera_Codigo)) > 0 Then
'         p_Mensaje.RecipIndex = r_int_Contad
'         p_Mensaje.RecipDisplayName = p_Arregl(r_int_Contad + 1).Genera_Codigo
'      End If
'   Next r_int_Contad
'
'   p_Mensaje.MsgSubject = p_Asunto
'   p_Mensaje.MsgNoteText = p_Contenido
'
'   r_int_Index = 0
'   If Len(Trim(txt_Adjunto1.Text)) > 0 Then
'      p_Mensaje.AttachmentIndex = r_int_Index
'      p_Mensaje.AttachmentName = txt_Adjunto1.Text 'p_NomFil_01
'      p_Mensaje.AttachmentPathName = txt_Adjunto1.Tag 'p_RutFil_01
'      p_Mensaje.AttachmentPosition = r_int_Index
'      p_Mensaje.AttachmentType = mapData
'   End If
'
'   If Len(Trim(txt_Adjunto2.Text)) > 0 Then
'      r_int_Index = r_int_Index + 1
'      p_Mensaje.AttachmentIndex = r_int_Index
'      p_Mensaje.AttachmentName = txt_Adjunto2.Text 'p_NomFil_02
'      p_Mensaje.AttachmentPathName = txt_Adjunto2.Tag 'p_RutFil_02
'      p_Mensaje.AttachmentPosition = r_int_Index
'      p_Mensaje.AttachmentType = mapData
'   End If
'
'   If Len(Trim(txt_Adjunto3.Text)) > 0 Then
'      r_int_Index = r_int_Index + 1
'      p_Mensaje.AttachmentIndex = r_int_Index
'      p_Mensaje.AttachmentName = txt_Adjunto3.Text 'p_NomFil_02
'      p_Mensaje.AttachmentPathName = txt_Adjunto3.Tag 'p_RutFil_02
'      p_Mensaje.AttachmentPosition = r_int_Index
'      p_Mensaje.AttachmentType = mapData
'   End If
'
'   p_Mensaje.Send
'   DoEvents
'   fs_EnviarCorreo = ""
'
'  'Cierra la sesión
'  p_Sesion.SignOff
'  Exit Function
'
'moddat_gf_EnvCor:
'   fs_EnviarCorreo = Err.Description
'   p_Sesion.SignOff
'   MsgBox Err.Description, vbCritical
'End Function

Private Sub cmd_Previa1_Click()
Dim res As Variant

   If Trim(txt_Adjunto1.Text) <> "" Then
      res = ShellExecute(1, "Open", txt_Adjunto1.Tag, "", "", 1)
   End If
End Sub

Private Sub cmd_Previa2_Click()
Dim res As Variant

   If Trim(txt_Adjunto2.Text) <> "" Then
      res = ShellExecute(1, "Open", txt_Adjunto2.Tag, "", "", 1)
   End If
End Sub

Private Sub cmd_Previa3_Click()
Dim res As Variant

   If Trim(txt_Adjunto3.Text) <> "" Then
      res = ShellExecute(1, "Open", txt_Adjunto3.Tag, "", "", 1)
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)

   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
      
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_str_Parame   As String
Dim r_rst_Genera   As ADODB.Recordset
Dim r_str_Mensaj   As String
   
   ReDim l_arr_DatEmp(0)
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT SEGEMP_CODIGO, SEGEMP_RAZSOC, DATEMP_DIRELE1, "
   r_str_Parame = r_str_Parame & "       DATEMP_DIRELE2 , DATEMP_DIRELE3, DATEMP_DIRELE4, DATEMP_DIRELE5 "
   r_str_Parame = r_str_Parame & "  FROM MNT_SEGEMP A "
   r_str_Parame = r_str_Parame & "  LEFT JOIN MNT_DATEMP B ON A.SEGEMP_CODIGO = B.DATEMP_CODEMP AND DATEMP_TIPTAB = 3 "
   r_str_Parame = r_str_Parame & " ORDER BY SEGEMP_RAZSOC ASC "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
     r_rst_Genera.Close
     Set r_rst_Genera = Nothing
     Exit Sub
   End If
   
   r_rst_Genera.MoveFirst
   Do While Not r_rst_Genera.EOF
      cmb_EmpSeg.AddItem Trim(r_rst_Genera!SEGEMP_RAZSOC)
      cmb_EmpSeg.ItemData(cmb_EmpSeg.NewIndex) = r_rst_Genera!SEGEMP_CODIGO
      
      '***AGREGAR AL ARREGLO
      ReDim Preserve l_arr_DatEmp(UBound(l_arr_DatEmp) + 1)
      l_arr_DatEmp(UBound(l_arr_DatEmp)).regEmp_CodEmp = Trim(r_rst_Genera!SEGEMP_CODIGO & "")
      l_arr_DatEmp(UBound(l_arr_DatEmp)).regEmp_Direle1 = Trim(r_rst_Genera!DATEMP_DIRELE1 & "")
      l_arr_DatEmp(UBound(l_arr_DatEmp)).regEmp_Direle2 = Trim(r_rst_Genera!DATEMP_DIRELE2 & "")
      l_arr_DatEmp(UBound(l_arr_DatEmp)).regEmp_Direle3 = Trim(r_rst_Genera!DATEMP_DIRELE3 & "")
      l_arr_DatEmp(UBound(l_arr_DatEmp)).regEmp_Direle4 = Trim(r_rst_Genera!DATEMP_DIRELE4 & "")
      l_arr_DatEmp(UBound(l_arr_DatEmp)).regEmp_Direle5 = Trim(r_rst_Genera!DATEMP_DIRELE5 & "")
            
      r_rst_Genera.MoveNext
   Loop
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
   '************************************************************
      
   txt_Destino.Enabled = False
   txt_Adjunto1.Enabled = False
   txt_Adjunto2.Enabled = False
   txt_Adjunto3.Enabled = False
   
   cmb_EmpSeg.ListIndex = -1
   
   txt_Destino.Text = ""
   txt_Asunto.Text = ""
   txt_Adjunto1.Text = ""
   txt_Adjunto1.Tag = ""
   txt_Adjunto2.Text = ""
   txt_Adjunto2.Tag = ""
   txt_Adjunto3.Text = ""
   txt_Adjunto3.Tag = ""
   txt_Mensaje.Text = ""
   
   txt_Asunto.Text = "DESGRAVAMEN EDPYME MICASITA - (" & Trim(pnl_Client.Caption) & ")"

   r_str_Mensaj = ""
   r_str_Mensaj = r_str_Mensaj & "Estimado(a)," & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "Adjunto las DPS del siguiente cliente para su evaluación:" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "                   1: " & moddat_g_str_NomCli & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "Quedamos a la espera de su respuesta." & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "Saludos cordiales," & vbCrLf
                                                                                                                                                                                                                                                       
   txt_Mensaje.Text = r_str_Mensaj
End Sub

Private Sub txt_Asunto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Mensaje)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Mensaje_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub
