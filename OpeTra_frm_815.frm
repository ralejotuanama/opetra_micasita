VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Con_PrePgo_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   Icon            =   "OpeTra_frm_815.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   4605
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   8123
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
         Left            =   60
         TabIndex        =   4
         Top             =   780
         Width           =   6495
         _Version        =   65536
         _ExtentX        =   11456
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
            Left            =   5880
            Picture         =   "OpeTra_frm_815.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprimir 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_815.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   6495
         _Version        =   65536
         _ExtentX        =   11456
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
            Left            =   660
            TabIndex        =   7
            Top             =   120
            Width           =   4665
            _Version        =   65536
            _ExtentX        =   8229
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Reporte de Conformidad - Prepago Parcial"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   6000
            Top             =   120
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
            Left            =   80
            Picture         =   "OpeTra_frm_815.frx":0890
            Top             =   80
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1215
         Left            =   60
         TabIndex        =   8
         Top             =   2040
         Width           =   6495
         _Version        =   65536
         _ExtentX        =   11456
         _ExtentY        =   2143
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
         Begin Threed.SSPanel pnl_Nombre_Tit 
            Height          =   315
            Left            =   1200
            TabIndex        =   9
            Top             =   420
            Width           =   5175
            _Version        =   65536
            _ExtentX        =   9128
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
         Begin EditLib.fpDoubleSingle txt_Monto_Tit 
            Height          =   315
            Left            =   1200
            TabIndex        =   0
            Top             =   765
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
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
            ThreeDInsideHighlightColor=   -2147483637
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
            ThreeDTextHighlightColor=   -2147483637
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Titular"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Monto:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   825
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   495
            Width           =   600
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   525
         Left            =   60
         TabIndex        =   12
         Top             =   1470
         Width           =   6495
         _Version        =   65536
         _ExtentX        =   11456
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
         Begin Threed.SSPanel pnl_ValorComercial 
            Height          =   315
            Left            =   4440
            TabIndex        =   13
            Top             =   105
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
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
         Begin Threed.SSPanel pnl_TipoPrepago 
            Height          =   315
            Left            =   1200
            TabIndex        =   14
            Top             =   105
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Valor Comercial:"
            Height          =   195
            Left            =   3120
            TabIndex        =   16
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Prepago:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   180
            Width           =   1005
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1215
         Left            =   60
         TabIndex        =   17
         Top             =   3315
         Width           =   6495
         _Version        =   65536
         _ExtentX        =   11456
         _ExtentY        =   2143
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
         Begin Threed.SSPanel pnl_Nombre_cony 
            Height          =   315
            Left            =   1200
            TabIndex        =   18
            Top             =   450
            Width           =   5175
            _Version        =   65536
            _ExtentX        =   9128
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
         Begin EditLib.fpDoubleSingle txt_Monto_Cony 
            Height          =   315
            Left            =   1200
            TabIndex        =   1
            Top             =   795
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
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
            ThreeDInsideHighlightColor=   -2147483637
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
            ThreeDTextHighlightColor=   -2147483637
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label8 
            Caption         =   "Conyuge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   525
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Monto:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   855
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frm_Con_PrePgo_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Imprimir_Click()
Dim r_str_Cadena As String

    r_str_Cadena = ""
    Call moddat_gs_FecSis

   If (Trim(pnl_Nombre_Tit.Caption) = "" And Trim(pnl_Nombre_cony.Caption) = "") Then
      MsgBox "No hay cliente titular ni cónyuge para imprimir.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If (Trim(pnl_Nombre_Tit.Caption) <> "" And Trim(pnl_Nombre_cony.Caption) <> "") Then
      If (CDbl(txt_Monto_Tit.Text) = 0 And CDbl(txt_Monto_Cony.Text) = 0) Then
         MsgBox "Ingrese un monto al titular o al cónyuge.", vbInformation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Monto_Tit)
         Exit Sub
      End If
   End If
   If (Trim(pnl_Nombre_Tit.Caption) <> "" And Trim(pnl_Nombre_cony.Caption) = "") Then
      If (CDbl(txt_Monto_Tit.Text) = 0) Then
         MsgBox "Ingrese un monto al titular.", vbInformation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Monto_Tit)
         Exit Sub
      End If
   End If
   If (Trim(pnl_Nombre_Tit.Caption) = "" And Trim(pnl_Nombre_cony.Caption) <> "") Then
      If (CDbl(txt_Monto_Cony.Text) = 0) Then
         MsgBox "Ingrese un monto al cónyuge.", vbInformation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Monto_Cony)
         Exit Sub
      End If
   End If
        
   'Proceso
   Screen.MousePointer = 11
      
   crp_Imprim.ParameterFields(0) = "p_nomtitu;" & Trim(CStr(moddat_g_str_NomCli & " ")) & ";True"
   crp_Imprim.ParameterFields(1) = "p_numtitu;" & moddat_gf_Consulta_ParDes("203", CStr(moddat_g_int_TipDoc)) & " - " & Trim(moddat_g_str_NumDoc & "") & ";True"
   
   If (Trim(moddat_g_str_CodIte) = "1") Then
      crp_Imprim.ParameterFields(2) = "p_tipoplazo;" & " " & ";True"
      crp_Imprim.ParameterFields(3) = "p_tipocuota;" & "X" & ";True"
   Else
      crp_Imprim.ParameterFields(2) = "p_tipocuota;" & " " & ";True"
      crp_Imprim.ParameterFields(3) = "p_tipoplazo;" & "x" & ";True"
   End If
      
   crp_Imprim.ParameterFields(5) = "p_montotitu;" & gf_FormatoNumero(txt_Monto_Tit.Text, 12, 2) & ";True"
   crp_Imprim.ParameterFields(6) = "p_simbolo;" & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & ";True"
   crp_Imprim.ParameterFields(7) = "p_valorcomercial;" & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(pnl_ValorComercial.Caption, 12, 2) & ";True"
   If (moddat_g_int_TipMon = 1) Then
       crp_Imprim.ParameterFields(8) = "p_moneda;" & "Soles" & ";True"
       crp_Imprim.ParameterFields(9) = "p_numcuenta;" & "Cuenta Ahorro soles N° 0011- 0369-02-00090532" & ";True"
   Else
       crp_Imprim.ParameterFields(8) = "p_moneda;" & "Dolares" & ";True"
       crp_Imprim.ParameterFields(9) = "p_numcuenta;" & "Cuenta Ahorro dólares N° 0011- 0369-02-00090540" & ";True"
   End If
   crp_Imprim.ParameterFields(10) = "p_nomcony;" & Trim(CStr(moddat_g_str_CygNom & " ")) & ";True"
   crp_Imprim.ParameterFields(11) = "p_numcony;" & moddat_gf_Consulta_ParDes("203", CStr(moddat_g_int_CygTDo)) & " - " & Trim(moddat_g_str_CygNDo & "") & ";True"
   crp_Imprim.ParameterFields(12) = "p_montocony;" & gf_FormatoNumero(txt_Monto_Cony.Text, 12, 2) & ";True"
   crp_Imprim.ParameterFields(13) = "p_simbolo;" & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & ";True"

   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ope_prepago_01.rpt"
      
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   'Exportar a crystal Report, crptRTF
   crp_Imprim.PrintFileType = crptCrystal
   crp_Imprim.Destination = crptToFile
   r_str_Cadena = "_" & Format(moddat_g_str_FecSis, "yyyymmdd") & "_" & Format(moddat_g_str_HorSis, "hhmmss")
   crp_Imprim.PrintFileName = g_str_RutLog & "\AFP_CONFORMIDAD\" & moddat_g_str_NumOpe & "_" & modgen_g_str_CodUsu & r_str_Cadena & ".RPT"
   crp_Imprim.Action = 1
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_str_Cadena As String
   
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   r_str_Cadena = ""
   txt_Monto_Tit.Enabled = False
   txt_Monto_Cony.Enabled = False
   
   If (Len(Trim(CStr(moddat_g_str_NomCli & " "))) > 0) Then
       pnl_Nombre_Tit.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
       txt_Monto_Tit.Text = 0
       txt_Monto_Tit.Enabled = True
   End If
   If (Len(Trim(CStr(moddat_g_str_CygNom & " "))) > 0) Then
       pnl_Nombre_cony.Caption = CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom
       txt_Monto_Cony.Text = 0
       txt_Monto_Cony.Enabled = True
   End If
   
   pnl_TipoPrepago.Caption = Trim(moddat_g_str_TipPar)
   pnl_ValorComercial.Caption = "0.00"
   If (Len(Trim(moddat_g_str_DesObs)) > 0) Then
      pnl_ValorComercial.Caption = Format(moddat_g_str_DesObs, "###,##0.00")
   End If
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
    pnl_Nombre_Tit.Caption = ""
    txt_Monto_Tit.Text = 0
    pnl_Nombre_cony.Caption = ""
    txt_Monto_Cony.Text = 0
    pnl_ValorComercial.Caption = "0.00"
    pnl_TipoPrepago.Caption = ""
End Sub

Private Sub txt_Monto_Cony_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Imprimir)
   End If
End Sub

Private Sub txt_Monto_Tit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_Monto_Cony.Enabled = False) Then
          Call gs_SetFocus(cmd_Imprimir)
      Else
          Call gs_SetFocus(txt_Monto_Cony)
      End If
   End If
End Sub

