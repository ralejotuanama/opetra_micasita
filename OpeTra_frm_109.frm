VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_TraCof_15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   1950
   ClientTop       =   3990
   ClientWidth     =   11250
   Icon            =   "OpeTra_frm_109.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4020
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   7091
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
         Height          =   765
         Left            =   30
         TabIndex        =   7
         Top             =   2730
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
         Begin VB.ComboBox cmb_RepLg2 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   9255
         End
         Begin VB.ComboBox cmb_RepLg1 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   9255
         End
         Begin VB.Label Label6 
            Caption         =   "Rep. Legal 2:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   390
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Rep. Legal 1:"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   8
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
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   9255
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo de Formato:"
            Height          =   285
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   10
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
            Height          =   495
            Left            =   630
            TabIndex        =   11
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Trámites COFIDE"
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
            Left            =   10650
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
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_109.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   12
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
            TabIndex        =   13
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
            Left            =   9690
            TabIndex        =   14
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
            Left            =   1860
            TabIndex        =   15
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
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8040
            TabIndex        =   17
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   19
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_109.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir Formato"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "OpeTra_frm_109.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   22
         Top             =   3540
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
         Begin EditLib.fpDateTime ipp_FecDes 
            Height          =   315
            Left            =   1860
            TabIndex        =   3
            Top             =   60
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
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
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha Desembolso:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frm_TraCof_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_RepLg1()      As moddat_tpo_Genera
Dim l_arr_RepLg2()      As moddat_tpo_Genera
Dim l_dbl_MtoCre        As Double
Dim l_dbl_ApoPro        As Double
Dim l_dbl_ComVta        As Double
Dim l_int_NumCuo        As Integer
Dim l_int_PerGra        As Integer
Dim l_int_CuoTot        As Integer
Dim l_dbl_IntGra        As Double
Dim l_dbl_CuoMen        As Double
Dim l_dbl_TasInt        As Double
Dim l_int_TipMon        As Integer
Dim l_str_FlgBTe        As String
Dim l_str_FlgBFu        As String
Dim l_str_FirCon        As String
Dim l_str_FirCvt        As String
Dim l_dbl_TCaApl        As Double
Dim l_dbl_MtoHip        As Double
Dim l_str_Direcc        As String
Dim l_str_IntDpt        As String
Dim l_str_NomZon        As String
Dim l_str_Depart        As String
Dim l_str_Provin        As String
Dim l_str_Distri        As String
Dim l_str_FlgCas        As String
Dim l_str_FlgDpt        As String
Dim l_str_FecApr        As String
Dim l_str_ClaSbs        As String
Dim l_str_ClaMCs        As String
Dim l_dbl_TCaSBS        As Double

Private Sub cmb_RepLg1_Click()
   Call gs_SetFocus(cmb_RepLg2)
End Sub

Private Sub cmb_RepLg1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_RepLg1_Click
   End If
End Sub

Private Sub cmb_RepLg2_Click()
   If ipp_FecDes.Enabled Then
      Call gs_SetFocus(ipp_FecDes)
   Else
      Call gs_SetFocus(cmd_Imprim)
   End If
End Sub

Private Sub cmb_RepLg2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_RepLg2_Click
   End If
End Sub

Private Sub cmb_TipRep_Click()
   If cmb_TipRep.ListIndex > -1 Then
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 11 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 13 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 21 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 22 Then
         cmb_RepLg1.Enabled = True
         cmb_RepLg2.Enabled = True
         
         ipp_FecDes.Enabled = False
         
         Call gs_SetFocus(cmb_RepLg1)
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 12 Then
         cmb_RepLg1.Enabled = False
         cmb_RepLg2.Enabled = False
         
         ipp_FecDes.Enabled = True
         
         Call gs_SetFocus(ipp_FecDes)
      Else
         cmb_RepLg1.ListIndex = -1
         cmb_RepLg2.ListIndex = -1
         
         ipp_FecDes.Text = Format(Date, "dd/mm/yyyy")
         
         cmb_RepLg1.Enabled = False
         cmb_RepLg2.Enabled = False
         
         ipp_FecDes.Enabled = False
         
         Call gs_SetFocus(cmd_Imprim)
      End If
   End If
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipRep_Click
   End If
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_str_ParEnt     As String
   Dim r_str_ParDec     As String
   
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Formato.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If

   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 11 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 13 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 21 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 22 Then
      If cmb_RepLg1.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Representante Legal que firmará el formato.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_RepLg1)
         Exit Sub
      End If
   
      If cmb_RepLg2.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Representante Legal que firmará el formato.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_RepLg2)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de imprimir el Formato?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_FORCOF WHERE "
   g_str_Parame = g_str_Parame & "FORCOF_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If

   'Grabando datos en Tabla Temporal
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "INSERT INTO RPT_FORCOF("
   g_str_Parame = g_str_Parame & "FORCOF_NUMSOL, "
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 11 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 21 Then
      g_str_Parame = g_str_Parame & "FORCOF_REPLG1, "
      g_str_Parame = g_str_Parame & "FORCOF_REPLG2) "
   End If
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 13 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 22 Then
      g_str_Parame = g_str_Parame & "FORCOF_REPLG1, "
      g_str_Parame = g_str_Parame & "FORCOF_REPLG2, "
      g_str_Parame = g_str_Parame & "FORCOF_MTOLET) "
   End If
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 12 Then
      g_str_Parame = g_str_Parame & "FORCOF_SUCURS, "
      g_str_Parame = g_str_Parame & "FORCOF_NOMCLI, "
      g_str_Parame = g_str_Parame & "FORCOF_DOCIDE, "
      g_str_Parame = g_str_Parame & "FORCOF_MTOPRE, "
      g_str_Parame = g_str_Parame & "FORCOF_NUMCUO, "
      g_str_Parame = g_str_Parame & "FORCOF_APOINI, "
      g_str_Parame = g_str_Parame & "FORCOF_PERGRA, "
      g_str_Parame = g_str_Parame & "FORCOF_VALVIV, "
      g_str_Parame = g_str_Parame & "FORCOF_FECCON, "
      g_str_Parame = g_str_Parame & "FORCOF_FECDES, "
      g_str_Parame = g_str_Parame & "FORCOF_MODTER, "
      g_str_Parame = g_str_Parame & "FORCOF_MODFUT, "
      g_str_Parame = g_str_Parame & "FORCOF_DIRECC, "
      g_str_Parame = g_str_Parame & "FORCOF_INTDPT, "
      g_str_Parame = g_str_Parame & "FORCOF_URBANI, "
      g_str_Parame = g_str_Parame & "FORCOF_DISTRI, "
      g_str_Parame = g_str_Parame & "FORCOF_PROVIN, "
      g_str_Parame = g_str_Parame & "FORCOF_DEPART, "
      g_str_Parame = g_str_Parame & "FORCOF_MTOHIP, "
      g_str_Parame = g_str_Parame & "FORCOF_FECMIN, "
      g_str_Parame = g_str_Parame & "FORCOF_TCAAPL) "
   End If
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 14 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 23 Then
      g_str_Parame = g_str_Parame & "FORCOF_SUCURS, "
      g_str_Parame = g_str_Parame & "FORCOF_SUCDPT, "
      g_str_Parame = g_str_Parame & "FORCOF_SUCPRV, "
      g_str_Parame = g_str_Parame & "FORCOF_SUCDST, "
      g_str_Parame = g_str_Parame & "FORCOF_FECAPR, "
      g_str_Parame = g_str_Parame & "FORCOF_NOMCLI, "
      g_str_Parame = g_str_Parame & "FORCOF_DOCIDE, "
      g_str_Parame = g_str_Parame & "FORCOF_CLASBS, "
      g_str_Parame = g_str_Parame & "FORCOF_CLAINT, "
      g_str_Parame = g_str_Parame & "FORCOF_MTOPRE, "
      g_str_Parame = g_str_Parame & "FORCOF_NUMCUO, "
      g_str_Parame = g_str_Parame & "FORCOF_APOINI, "
      g_str_Parame = g_str_Parame & "FORCOF_PERGRA, "
      g_str_Parame = g_str_Parame & "FORCOF_INTGRA, "
      g_str_Parame = g_str_Parame & "FORCOF_CUOFIJ, "
      g_str_Parame = g_str_Parame & "FORCOF_TASINT, "
      g_str_Parame = g_str_Parame & "FORCOF_VALVIV, "
      g_str_Parame = g_str_Parame & "FORCOF_FECCON, "
      g_str_Parame = g_str_Parame & "FORCOF_MODTER, "
      g_str_Parame = g_str_Parame & "FORCOF_MODFUT, "
      g_str_Parame = g_str_Parame & "FORCOF_DIRECC, "
      g_str_Parame = g_str_Parame & "FORCOF_INTDPT, "
      g_str_Parame = g_str_Parame & "FORCOF_URBANI, "
      g_str_Parame = g_str_Parame & "FORCOF_DISTRI, "
      g_str_Parame = g_str_Parame & "FORCOF_PROVIN, "
      g_str_Parame = g_str_Parame & "FORCOF_DEPART, "
      g_str_Parame = g_str_Parame & "FORCOF_MTOHIP, "
      g_str_Parame = g_str_Parame & "FORCOF_FECMIN, "
      g_str_Parame = g_str_Parame & "FORCOF_TCAAPL, "
      g_str_Parame = g_str_Parame & "FORCOF_INMCAS, "
      g_str_Parame = g_str_Parame & "FORCOF_INMDPT, "
      g_str_Parame = g_str_Parame & "FORCOF_TCASBS) "
   End If
   
   g_str_Parame = g_str_Parame & "VALUES ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 11 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 21 Then
      g_str_Parame = g_str_Parame & "'" & cmb_RepLg1.Text & "', "
      g_str_Parame = g_str_Parame & "'" & cmb_RepLg2.Text & "') "
   End If
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 13 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 22 Then
      r_str_ParEnt = gf_Convertir_NumLet(CLng(Mid(Format(l_dbl_MtoCre, "###,##0.00"), 1, InStr(Format(l_dbl_MtoCre, "###,##0.00"), ".") - 1)))
   
      r_str_ParDec = Right(Format(l_dbl_MtoCre, "###,##0.00"), 2)
   
      g_str_Parame = g_str_Parame & "'" & cmb_RepLg1.Text & "', "
      g_str_Parame = g_str_Parame & "'" & cmb_RepLg2.Text & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_ParEnt & " y " & r_str_ParDec & "/100 " & moddat_gf_Consulta_ParDes("204", CStr(l_int_TipMon)) & "') "
   End If
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 12 Then
      g_str_Parame = g_str_Parame & "'PRINCIPAL', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NomCli & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoCre) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ApoPro) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_PerGra) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ComVta) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_FirCon & "', "
      g_str_Parame = g_str_Parame & "'" & ipp_FecDes.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_FlgBTe & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_FlgBFu & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Direcc & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_IntDpt & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Distri & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Provin & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Depart & "', "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoHip) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_FirCvt & "', "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TCaApl) & ") "
   End If
   
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 14 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 23 Then
      g_str_Parame = g_str_Parame & "'PRINCIPAL', "
      
      g_str_Parame = g_str_Parame & "'LIMA', "
      g_str_Parame = g_str_Parame & "'LIMA', "
      g_str_Parame = g_str_Parame & "'SAN ISIDRO', "
      g_str_Parame = g_str_Parame & "'" & l_str_FecApr & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NomCli & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_ClaSbs & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_ClaMCs & "', "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoCre) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ApoPro) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_PerGra) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IntGra) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CuoMen) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TasInt) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ComVta) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_FirCon & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_FlgBTe & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_FlgBFu & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Direcc & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_IntDpt & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Distri & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Provin & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Depart & "', "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoHip) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_FirCvt & "', "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TCaApl) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_FlgCas & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_FlgDpt & "',"
      g_str_Parame = g_str_Parame & CStr(l_dbl_TCaSBS) & ") "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If

   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 62, 38, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = UCase(moddat_g_str_EntDat) & ".CLI_DATGEN"
   crp_Imprim.DataFiles(2) = UCase(moddat_g_str_EntDat) & ".RPT_FORCOF"
   
   crp_Imprim.SelectionFormula = "{CRE_SOLMAE.SOLMAE_NUMERO} = '" & moddat_g_str_NumSol & "' "
   
   Select Case cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
      Case 11: crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_FORCOF_01.RPT"
      Case 12: crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_FORCOF_04.RPT"
      Case 13: crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_FORCOF_02.RPT"
      Case 14: crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_FORCOF_05.RPT"
      Case 21: crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_FORCOF_03.RPT"
      Case 22: crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_FORCOF_07.RPT"
      Case 23: crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_FORCOF_06.RPT"
   End Select
   
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_DatCre
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_RepLg1, l_arr_RepLg1, 1, "512")
   Call moddat_gs_Carga_LisIte(cmb_RepLg2, l_arr_RepLg2, 1, "512")

   cmb_TipRep.Clear
   If moddat_g_str_CodPrd = "003" Then
      cmb_TipRep.AddItem "ANEXO A - CARTA SOLICITUD DE DESEMBOLSO"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 11
      
      cmb_TipRep.AddItem "ANEXO B - INFORME DE CREDITO"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 12
      
      cmb_TipRep.AddItem "ANEXO C - PAGARE"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 13
      
      cmb_TipRep.AddItem "ANEXO D - SOLICITUD DE CRC Y PBP"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 14
   Else
      cmb_TipRep.AddItem "ANEXO A - CARTA SOLICITUD DE DESEMBOLSO"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 21
   
      cmb_TipRep.AddItem "ANEXO B - PAGARE"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 22
   
      cmb_TipRep.AddItem "ANEXO C - INFORME DE CREDITO"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 23
   End If
   
   'Inicializando Controles
   cmb_TipRep.ListIndex = -1
   cmb_RepLg1.ListIndex = -1
   cmb_RepLg2.ListIndex = -1
   
   ipp_FecDes.Text = Format(Date, "dd/mm/yyyy")
   
   cmb_RepLg1.Enabled = False
   cmb_RepLg2.Enabled = False
   
   ipp_FecDes.Enabled = False
End Sub

Private Sub fs_DatCre()
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   l_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   
   Select Case g_rst_Princi!SOLMAE_TIPMON
      Case 1
         l_dbl_MtoCre = g_rst_Princi!SOLMAE_MTOPRE_SOL
         l_dbl_ApoPro = g_rst_Princi!SOLMAE_APOPRO_SOL
         l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_SOL
         
      Case 2
         l_dbl_MtoCre = g_rst_Princi!SOLMAE_MTOPRE_DOL
         l_dbl_ApoPro = g_rst_Princi!SOLMAE_APOPRO_DOL
         l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_DOL
   End Select
   
   l_int_NumCuo = g_rst_Princi!SOLMAE_NUMCUO
   l_int_PerGra = g_rst_Princi!SOLMAE_PERGRA
   l_int_CuoTot = g_rst_Princi!SOLMAE_NUMCUO + g_rst_Princi!SOLMAE_PERGRA
   l_dbl_IntGra = g_rst_Princi!SOLMAE_INTGRA
   
   l_dbl_CuoMen = g_rst_Princi!SOLMAE_CUOAPR_MPR
   l_dbl_TasInt = g_rst_Princi!SOLMAE_TASINT
   
   l_str_FlgBTe = " "
   l_str_FlgBFu = " "
   Select Case CInt(g_rst_Princi!SOLMAE_CODMOD)
      Case 1:  l_str_FlgBTe = "x"
      Case 2:  l_str_FlgBFu = "x"
      Case 3:  l_str_FlgBFu = "x"
   End Select
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Datos de Legal
   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      l_str_FirCon = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
      l_str_FirCvt = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCVT))
      l_dbl_TCaSBS = g_rst_Princi!EVALEG_TCASBS
      l_dbl_TCaApl = g_rst_Princi!EVALEG_TCACVT
      l_dbl_MtoHip = g_rst_Princi!EVALEG_MTOHIP
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Datos del Inmueble
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      l_str_Direcc = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & " "
      l_str_Direcc = l_str_Direcc & Trim(g_rst_Princi!SOLINM_NOMVIA & "") & " " & Trim(g_rst_Princi!SOLINM_NUMVIA & "") & IIf(Len(Trim(g_rst_Princi!SOLINM_ESTACI)) = 0 Or Trim(g_rst_Princi!SOLINM_ESTACI) = "NO", "", " - (ESTACIONAMIENTO: " & Trim(g_rst_Princi!SOLINM_ESTACI) & ")")
      
      l_str_IntDpt = Trim(g_rst_Princi!SOLINM_INTDPT & "")
      
      Select Case g_rst_Princi!SOLINM_TIPZON
         Case 1
            l_str_NomZon = Trim(g_rst_Princi!SOLINM_NOMZON & "")
      End Select
      
      l_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000")
      l_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00")
      l_str_Distri = moddat_gf_Consulta_ParDes("101", g_rst_Princi!SOLINM_UBIGEO)
      
      If g_rst_Princi!SOLINM_TIPINM = 1 Then
         l_str_FlgCas = "x"
      ElseIf g_rst_Princi!SOLINM_TIPINM = 2 Then
         l_str_FlgDpt = "x"
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Fecha de Aprobación Crediticia
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 21"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   g_rst_Genera.MoveFirst

   If g_rst_Genera!SEGUIM_SITUAC = 1 Then
      l_str_FecApr = gf_FormatoFecha(CStr(g_rst_Genera!SEGUIM_FECFIN))
   End If

   'Datos del Cliente
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      l_str_ClaSbs = Trim(g_rst_Princi!DATGEN_CLASBS & "")
      l_str_ClaMCs = Trim(g_rst_Princi!DATGEN_CLASMC & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub


