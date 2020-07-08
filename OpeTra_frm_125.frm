VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_ModSol_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2745
   ClientLeft      =   4125
   ClientTop       =   4095
   ClientWidth     =   11625
   Icon            =   "OpeTra_frm_125.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2745
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   4
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
            Picture         =   "OpeTra_frm_125.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_125.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   5
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
            Left            =   630
            TabIndex        =   6
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10920
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
            Left            =   10290
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   630
            TabIndex        =   7
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cambio de Tasa de Interés"
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
            Picture         =   "OpeTra_frm_125.frx":0890
            Top             =   60
            Width           =   480
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
      Begin Threed.SSPanel SSPanel11 
         Height          =   435
         Left            =   30
         TabIndex        =   13
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
         Begin EditLib.fpDoubleSingle ipp_TasInt 
            Height          =   315
            Left            =   1440
            TabIndex        =   0
            Top             =   60
            Width           =   1275
            _Version        =   196608
            _ExtentX        =   2249
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.000000"
            DecimalPlaces   =   6
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label5 
            Caption         =   "Tasa de Interés:"
            Height          =   285
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frm_ModSol_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Lista de variables para ser usadas en el proceso de Auditoria
Dim l_dbl_TasInt  As Double

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

   r_str_Proceso = "CAMBIO TASA DE INTERES"
   r_str_Tabla = "CRE_SOLMAE"
   r_str_Usuario = modgen_g_str_CodUsu
   r_str_Terminal = modgen_g_str_NombPC
   r_str_Plataforma = UCase(App.EXEName)
   r_str_Sucursal = modgen_g_str_CodSuc

   r_str_Descri1 = ""
   r_str_Descri2 = ""
   r_str_Descri3 = ""

   'Verificacion de datos modificados para ser guardados como Auditoria
   If l_dbl_TasInt <> ipp_TasInt.Text Then
      r_str_Descri = r_str_Descri + "Tasa de Interes (Antes: " & Format(l_dbl_TasInt, "#,##0.000000") & ")  (Nuevo: " & ipp_TasInt.Text & ")" + Chr(13)
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
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_dbl_TasInt = ipp_TasInt.Value
   moddat_g_int_FlgAct_1 = 2
   Call Grabar_Auditoria
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
   ipp_TasInt.Value = moddat_g_dbl_TasInt
   
   'Asignacion de variables usadas en el Proceso de Auditoria.
   l_dbl_TasInt = ipp_TasInt.Value
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(ipp_TasInt)
   Screen.MousePointer = 0
End Sub

Private Sub ipp_TasInt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub
