VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Tra_EvaTas_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   Icon            =   "OpeTra_frm_354.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7650
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   13494
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
         Height          =   5340
         Left            =   30
         TabIndex        =   11
         Top             =   2250
         Width           =   10845
         _Version        =   65536
         _ExtentX        =   19129
         _ExtentY        =   9419
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
            TabIndex        =   35
            Text            =   "OpeTra_frm_354.frx":000C
            Top             =   2220
            Width           =   9255
         End
         Begin VB.TextBox txt_Adjunto3 
            Height          =   315
            Left            =   1470
            MaxLength       =   150
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1860
            Width           =   4850
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
            TabIndex        =   4
            Top             =   1860
            Width           =   525
         End
         Begin VB.CommandButton cmd_Previa3 
            Caption         =   "Ver"
            Height          =   315
            Left            =   10200
            TabIndex        =   32
            Top             =   1860
            Width           =   525
         End
         Begin VB.CommandButton cmd_Previa2 
            Caption         =   "Ver"
            Height          =   315
            Left            =   10200
            TabIndex        =   31
            Top             =   1500
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
            TabIndex        =   3
            Top             =   1500
            Width           =   525
         End
         Begin VB.CommandButton cmd_Previa1 
            Caption         =   "Ver"
            Height          =   315
            Left            =   10200
            TabIndex        =   30
            Top             =   1140
            Width           =   525
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
            TabIndex        =   2
            Top             =   1140
            Width           =   525
         End
         Begin VB.TextBox txt_Adjunto2 
            Height          =   315
            Left            =   1470
            MaxLength       =   150
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1500
            Width           =   4850
         End
         Begin VB.TextBox txt_Destino 
            Height          =   630
            Left            =   1470
            MaxLength       =   600
            MultiLine       =   -1  'True
            TabIndex        =   0
            Text            =   "OpeTra_frm_354.frx":0012
            Top             =   90
            Width           =   9255
         End
         Begin VB.TextBox txt_Adjunto1 
            Height          =   315
            Left            =   1470
            MaxLength       =   150
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1140
            Width           =   4850
         End
         Begin VB.TextBox txt_Asunto 
            Height          =   315
            Left            =   1470
            MaxLength       =   400
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   780
            Width           =   9255
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   6360
            TabIndex        =   27
            Top             =   1140
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "<= Reporte de Orden de Trabajo"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   315
            Left            =   6360
            TabIndex        =   28
            Top             =   1500
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "<= Minuta Compra / Venta"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   315
            Left            =   6360
            TabIndex        =   33
            Top             =   1860
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Adjunto 3:"
            Height          =   195
            Left            =   180
            TabIndex        =   34
            Top             =   1920
            Width           =   720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Adjunto 2:"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   1560
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Mensaje:"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   3420
            Width           =   645
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Adjunto 1:"
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Asunto:"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Destino:"
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   300
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   15
            Top             =   330
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Envio Correo de Orden de Trabajo"
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
            Left            =   10350
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
            Left            =   9720
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
            Left            =   9150
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   8670
            Top             =   30
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_354.frx":0018
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   16
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
            Left            =   9300
            TabIndex        =   18
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
            Left            =   1470
            TabIndex        =   19
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
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   390
            Width           =   525
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   7650
            TabIndex        =   21
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Solicitud"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   60
            Width           =   945
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   23
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
            Picture         =   "OpeTra_frm_354.frx":0322
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_EnvMail 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_354.frx":0764
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Envio de Correo"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_EvaTas_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Adjun2_Click()
  On Error GoTo cmd_BusArc_Error
 
   'dlg_Guarda.Filter = "Todos los archivos (*.*)|*.*"
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

Private Sub cmd_EnvMail_Click()
Dim r_str_Parame   As String
Dim r_rst_Princi   As ADODB.Recordset
Dim r_str_ResExt   As String
Dim r_int_PosExt   As Integer
Dim r_int_NumLOG   As Integer
Dim r_bol_Mensaj   As String
Dim r_str_NomAch   As String
Dim r_str_NumFil   As Integer

   If Trim(txt_Destino.Text) = "" Then
      MsgBox "Debe ingresar los correos de destino en el mantenedor de peritos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Asunto)
      Exit Sub
   End If
   
   If Trim(txt_Asunto.Text) = "" Then
      MsgBox "Debe ingresar un asunto para el envio de correo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Asunto)
      Exit Sub
   End If
   
   If Trim(txt_Adjunto1.Text) = "" Then
      MsgBox "No se genero el reporte de orden de trabajo - tasacion de inmueble.", vbExclamation, modgen_g_str_NomPlt
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
   
   Open g_str_RutLogTas & "\" & moddat_g_str_NumSol & "_41_0_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".LOG" For Output As r_int_NumLOG
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
         
   'Renombrando adjunto 2 y transfiriendolo a la carpeta de tasación
   r_str_ResExt = ""
   r_str_ResExt = fs_Extension(txt_Adjunto2.Text)
   If Len(Trim(txt_Adjunto2.Tag)) > 0 Then
      r_str_NomAch = moddat_g_str_NumSol & "_41_2_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & "." & r_str_ResExt
      
      FileCopy txt_Adjunto2.Tag, g_str_RutLogTas & r_str_NomAch
      Print #r_int_NumLOG, "*:     Transfiriendo el Adjunto_2(" & txt_Adjunto2.Tag & ") a la carpeta de tasacion, renombrandolo."
      Print #r_int_NumLOG, "       Nuevo nombre del archivo: " & g_str_RutLogTas & r_str_NomAch
   Else
      Print #r_int_NumLOG, "*:     No hay ningún archivo a adjuntar en el correo en Adjuntar_2"
   End If
            
   'Renombrando adjunto 3 y transfiriendolo a la carpeta de tasación
   r_str_ResExt = ""
   r_str_ResExt = fs_Extension(txt_Adjunto3.Text)
   If Len(Trim(txt_Adjunto3.Tag)) > 0 Then
      r_str_NomAch = moddat_g_str_NumSol & "_41_3_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & "." & r_str_ResExt
      
      FileCopy txt_Adjunto3.Tag, g_str_RutLogTas & r_str_NomAch
      Print #r_int_NumLOG, "*:     Transfiriendo el Adjunto_3(" & txt_Adjunto3.Tag & ") a la carpeta de tasacion, renombrandolo."
      Print #r_int_NumLOG, "       Nuevo nombre del archivo: " & g_str_RutLogTas & r_str_NomAch
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
   r_bol_Mensaj = ""
   r_bol_Mensaj = fs_EnviarCorreoAdj(mps_Sesion, mps_Mensaj, txt_Asunto.Text, txt_Mensaje.Text, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, False, False, False)
   
   If r_bol_Mensaj = "" Then
      Print #r_int_NumLOG, "*:     Se envio el correo correctamente."
   Else
      Print #r_int_NumLOG, "*:     " & r_bol_Mensaj
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If r_bol_Mensaj = "" Then
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 93, 0, "", 0, 0) Then
         Print #r_int_NumLOG, "*:     Error al insertar la ocurrencia en el seguimiento."
         Exit Sub
      Else
         Print #r_int_NumLOG, "*:     Se inserto la ocurrencia correctamente al seguimiento."
      End If
   End If
   
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, ""
   Print #r_int_NumLOG, "       Envio Correo de Orden de Trabajo "
   Print #r_int_NumLOG, "       ================================ "
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
   If r_bol_Mensaj = "" Then
      MsgBox "Se envió la orden de trabajo por correo.", vbInformation, modgen_g_str_NomPlt
   End If
   
   Screen.MousePointer = 0
   moddat_g_int_FlgAct_1 = 2
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
   
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_str_Parame    As String
Dim r_str_CadAux    As String
Dim r_str_NomPry    As String
Dim r_str_TipMod    As String
Dim r_rst_Princi    As ADODB.Recordset
Dim r_str_Mensaj    As String
    
   cmd_Adjun1.Enabled = False
   txt_Destino.Enabled = False
   txt_Adjunto1.Enabled = False
   txt_Adjunto2.Enabled = False
   txt_Adjunto3.Enabled = False
   
   txt_Destino.Text = ""
   txt_Asunto.Text = ""
   txt_Adjunto1.Text = ""
   txt_Adjunto1.Tag = ""
   
   txt_Mensaje.Text = ""
   txt_Adjunto2.Text = ""
   txt_Adjunto2.Tag = ""
   txt_Adjunto3.Text = ""
   txt_Adjunto3.Tag = ""
   r_str_CadAux = ""
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   '--------------------------------------
   txt_Asunto.Text = "ENVIO DE ORDEN DE TRABAJO (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"

   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT"
   r_str_Parame = r_str_Parame & "        (CASE WHEN H.SOLINM_TABPRY IS NOT NULL THEN"
   r_str_Parame = r_str_Parame & "             CASE WHEN H.SOLINM_TABPRY = 2 THEN"
   r_str_Parame = r_str_Parame & "                  CASE WHEN H.SOLINM_PRYCOD IS NOT NULL THEN"
   r_str_Parame = r_str_Parame & "                       CASE WHEN LENGTH (H.SOLINM_PRYCOD) > 0 THEN"
   r_str_Parame = r_str_Parame & "                            (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD)"
   r_str_Parame = r_str_Parame & "                       Else"
   r_str_Parame = r_str_Parame & "                            CASE WHEN LENGTH (H.SOLINM_PRYNOM) > 0 THEN TRIM(H.SOLINM_PRYNOM) END"
   r_str_Parame = r_str_Parame & "                        End"
   r_str_Parame = r_str_Parame & "                  Else"
   r_str_Parame = r_str_Parame & "                       CASE WHEN LENGTH (H.SOLINM_PRYCOD) > 0 THEN"
   r_str_Parame = r_str_Parame & "                            (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD)"
   r_str_Parame = r_str_Parame & "                       Else"
   r_str_Parame = r_str_Parame & "                            CASE WHEN H.SOLINM_PRYNOM IS NOT NULL THEN"
   r_str_Parame = r_str_Parame & "                              Trim (H.SOLINM_PRYNOM)"
   r_str_Parame = r_str_Parame & "                            Else ''"
   r_str_Parame = r_str_Parame & "                             End"
   r_str_Parame = r_str_Parame & "                        End"
   r_str_Parame = r_str_Parame & "                   End"
   r_str_Parame = r_str_Parame & "             Else"
   r_str_Parame = r_str_Parame & "                   CASE WHEN H.SOLINM_PRYCOD IS NOT NULL THEN"
   r_str_Parame = r_str_Parame & "                        (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD)"
   r_str_Parame = r_str_Parame & "                   Else"
   r_str_Parame = r_str_Parame & "                        CASE WHEN H.SOLINM_PRYNOM IS NOT NULL THEN"
   r_str_Parame = r_str_Parame & "                          Trim (H.SOLINM_PRYNOM)"
   r_str_Parame = r_str_Parame & "                        Else"
   r_str_Parame = r_str_Parame & "                          ''"
   r_str_Parame = r_str_Parame & "                         End"
   r_str_Parame = r_str_Parame & "                    End"
   r_str_Parame = r_str_Parame & "              End"
   r_str_Parame = r_str_Parame & "        Else"
   r_str_Parame = r_str_Parame & "              CASE WHEN H.SOLINM_PRYCOD IS NOT NULL THEN"
   r_str_Parame = r_str_Parame & "                (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD)"
   r_str_Parame = r_str_Parame & "              Else"
   r_str_Parame = r_str_Parame & "                ''"
   r_str_Parame = r_str_Parame & "               End"
   r_str_Parame = r_str_Parame & "         END) AS NOMBRE_PROYECTO,"
   r_str_Parame = r_str_Parame & "      (SELECT TRIM(PARPRD_DESCRI) "
   r_str_Parame = r_str_Parame & "         FROM CRE_PARPRD "
   r_str_Parame = r_str_Parame & "        WHERE PARPRD_CODPRD = SOLMAE_CODPRD "
   r_str_Parame = r_str_Parame & "          AND PARPRD_CODGRP = '003' "
   r_str_Parame = r_str_Parame & "          AND PARPRD_CODITE = LPAD(A.SOLMAE_CODMOD,3,'0') AND ROWNUM = 1) AS MODALIDAD "
   r_str_Parame = r_str_Parame & " FROM CRE_SOLMAE A "
   r_str_Parame = r_str_Parame & "INNER JOIN CRE_SOLINM H ON H.SOLINM_NUMSOL = A.SOLMAE_NUMERO "
   r_str_Parame = r_str_Parame & "WHERE H.SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   r_str_NomPry = "_______________"
   r_str_TipMod = "_______________"
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
         r_str_NomPry = Trim(r_rst_Princi!NOMBRE_PROYECTO & "")
         r_str_TipMod = IIf(InStr(1, Trim(r_rst_Princi!MODALIDAD & ""), "BIEN FUTURO") > 0, "BIEN FUTURO", "BIEN TERMINADO")
   End If
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
        
   '--------------------------------------
   r_str_Mensaj = ""
   r_str_Mensaj = r_str_Mensaj & "Estimados:" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "Adjunto la siguiente documentación para que se realice  la tasación correspondiente a  un '" & UCase(r_str_TipMod) & "'" & " del proyecto '" & UCase(Trim(r_str_NomPry)) & "', de los clientes en mención:" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "1. Orden de tasación." & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "2. Características del departamento." & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "Favor de coordinar la inspección: _________________________" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "confirmarla vía mail.  Se adjunta correo de referencia con documentos." & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "Quedamos a la espera del informe respectivo incluyendo las fotos del proyecto así como del C.02 del Fondo MIVIVIENDA; asimismo solicito puedan confirmarme la coordinación que se haga para la inspección del inmueble." & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "" & vbCrLf
   r_str_Mensaj = r_str_Mensaj & "Saludos cordiales"
   
   txt_Mensaje.Text = r_str_Mensaj
   '--------------------------------------
   
   txt_Adjunto1.Text = fs_GenExc
   If txt_Adjunto1.Text <> "" Then
      txt_Adjunto1.Tag = g_str_RutLogTas & txt_Adjunto1.Text
   End If
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT DATEMP_DIRELE1, DATEMP_DIRELE2, DATEMP_DIRELE3, DATEMP_DIRELE4, DATEMP_DIRELE5 "
   r_str_Parame = r_str_Parame & "   FROM MNT_DATEMP A "
   r_str_Parame = r_str_Parame & "  WHERE A.DATEMP_CODEMP = '" & moddat_g_str_CodGen & "'"
   r_str_Parame = r_str_Parame & "    AND A.DATEMP_TIPTAB = 1 "
   r_str_Parame = r_str_Parame & "    AND A.DATEMP_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      r_str_CadAux = ""
      
      If Trim(r_rst_Princi!DATEMP_DIRELE1 & "") <> "" Then
         r_str_CadAux = Trim(r_rst_Princi!DATEMP_DIRELE1 & "") & "; "
         ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
         moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim(r_rst_Princi!DATEMP_DIRELE1)
      End If
      
      If Trim(r_rst_Princi!DATEMP_DIRELE2 & "") <> "" Then
         r_str_CadAux = r_str_CadAux & Trim(r_rst_Princi!DATEMP_DIRELE2 & "") & "; "
         ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
         moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim(r_rst_Princi!DATEMP_DIRELE2)
      End If
      
      If Trim(r_rst_Princi!DATEMP_DIRELE3 & "") <> "" Then
         r_str_CadAux = r_str_CadAux & Trim(r_rst_Princi!DATEMP_DIRELE3 & "") & "; "
         ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
         moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim(r_rst_Princi!DATEMP_DIRELE3)
      End If
      
      If Trim(r_rst_Princi!DATEMP_DIRELE4 & "") <> "" Then
         r_str_CadAux = r_str_CadAux & Trim(r_rst_Princi!DATEMP_DIRELE4 & "") & "; "
         ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
         moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim(r_rst_Princi!DATEMP_DIRELE4)
      End If
      
      If Trim(r_rst_Princi!DATEMP_DIRELE5 & "") <> "" Then
         r_str_CadAux = r_str_CadAux & Trim(r_rst_Princi!DATEMP_DIRELE5 & "") & "; "
         ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
         moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim(r_rst_Princi!DATEMP_DIRELE5)
      End If
      
      txt_Destino.Text = r_str_CadAux
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   Call gs_SetFocus(txt_Asunto)
End Sub

'Private Function fs_EnviarCorreoAdj(p_Sesion As MAPISession, p_Mensaje As MAPIMessages, p_Asunto As String, p_Contenido As String, _
'                                    p_User1 As String, p_User2 As String, p_JfCred As Boolean, p_Legal As Boolean, p_DrAdm As Boolean) As String
'
'   Dim r_int_Contad      As Integer
'   Dim r_int_Index       As Integer
'
'   On Error GoTo moddat_gf_EnvCor
'
'   fs_EnviarCorreoAdj = ""
'   'Inicializa
'   p_Sesion.DownLoadMail = False
'   p_Sesion.NewSession = True
'   p_Sesion.SignOn
'   p_Mensaje.SessionID = p_Sesion.SessionID
'
'   'Envío
'   p_Mensaje.Compose
'   '-----------------------------------------------------------------------------------------------------
'
'   'Consejero Hipotecario
'   If (Len(Trim(p_User1)) > 0) Then
'       ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
'       moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User1))
'   End If
'
'   'Ejecutivo de Seguimientos
'   If (Len(Trim(p_User2)) > 0) Then
'       r_str_Cadena = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User2))
'       If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
'          ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
'          moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
'       End If
'   End If
'
'   'Director de Producción
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(200, moddat_g_arr_Genera)
'
'   'Director Comercial
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(100, moddat_g_arr_Genera)
'
'   'Jefe de Ventas
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(120, moddat_g_arr_Genera)
'
'   'Jefe de Seguimiento
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(130, moddat_g_arr_Genera)
'
'   'Jefe de Operaciones
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(220, moddat_g_arr_Genera)
'
'   'Evaluador de Operaciones
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(221, moddat_g_arr_Genera)
'
'   'Legal
'   If (p_Legal = True) Then
'      'Jefe de Legal
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(230, moddat_g_arr_Genera)
'
'      'Asistente de Legal 1
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(231, moddat_g_arr_Genera)
'
'      'Asistente de Legal 2
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(232, moddat_g_arr_Genera)
'   End If
'
'   'Creditos
'   If (p_JfCred = True) Then
'      'Jefe de Créditos
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(210, moddat_g_arr_Genera)
'
'      'Evaluadores de credito
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_UsuEje_Arr(211, moddat_g_arr_Genera, moddat_g_str_NumSol)
'   End If
'
'   If (p_DrAdm = True) Then
'      'Director de Administración
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(300, moddat_g_arr_Genera)
'   End If
'   '-----------------------------------------------------------------------------------------------------
'
'   For r_int_Contad = 0 To UBound(moddat_g_arr_Genera) - 1
'      If Len(Trim(moddat_g_arr_Genera(r_int_Contad + 1).Genera_Codigo)) > 0 Then
'         p_Mensaje.RecipIndex = r_int_Contad
'         p_Mensaje.RecipDisplayName = moddat_g_arr_Genera(r_int_Contad + 1).Genera_Codigo
'      End If
'   Next r_int_Contad
'
'   p_Mensaje.MsgSubject = p_Asunto
'   p_Mensaje.MsgNoteText = p_Contenido
'
'   For r_int_Contad = 0 To UBound(moddat_g_arr_GenAux) - 1
'      If Len(Trim(moddat_g_arr_GenAux(r_int_Contad + 1).Genera_Codigo)) > 0 Then
'         p_Mensaje.AttachmentIndex = r_int_Contad
'         p_Mensaje.AttachmentName = moddat_g_arr_GenAux(r_int_Contad + 1).Genera_Codigo  'p_NomFil_01
'         p_Mensaje.AttachmentPathName = moddat_g_arr_GenAux(r_int_Contad + 1).Genera_Refere 'p_RutFil_01
'         p_Mensaje.AttachmentPosition = r_int_Contad
'         p_Mensaje.AttachmentType = mapData
'      End If
'   Next r_int_Contad
'
'   'Enviar Correro
'   p_Mensaje.Send
'   DoEvents
'   fs_EnviarCorreoAdj = ""
'
'  'Cierra la sesión
'  p_Sesion.SignOff
'  Exit Function
'
'moddat_gf_EnvCor:
'   fs_EnviarCorreoAdj = Err.Description
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

Private Sub txt_Adjunto2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Asunto)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Asunto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Mensaje)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Adjunto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Mensaje)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
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
   'l_str_Percon_mail = Trim(r_rst_Princi!PERCON_DIRELE & "")
   
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
   
   'r_obj_Excel.ActiveWorkbook.SaveAs ("D:\PRUEBA.XLSX")
   
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

Private Sub txt_Mensaje_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub
