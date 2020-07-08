VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Pro_MViPag_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   2325
   ClientTop       =   2730
   ClientWidth     =   7875
   Icon            =   "OpeTra_frm_057.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4875
      Left            =   -30
      TabIndex        =   5
      Top             =   0
      Width           =   7905
      _Version        =   65536
      _ExtentX        =   13944
      _ExtentY        =   8599
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
         Height          =   765
         Left            =   30
         TabIndex        =   6
         Top             =   4050
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
         Begin VB.CommandButton cmd_Proces 
            Height          =   675
            Left            =   6420
            Picture         =   "OpeTra_frm_057.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7110
            Picture         =   "OpeTra_frm_057.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   930
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
            WindowShowRefreshBtn=   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3225
         Left            =   30
         TabIndex        =   7
         Top             =   780
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   5689
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
         Begin VB.DriveListBox drv_Listas 
            Height          =   315
            Left            =   1560
            TabIndex        =   13
            Top             =   720
            Width           =   6165
         End
         Begin VB.DirListBox dir_Carpet 
            Height          =   2115
            Left            =   1560
            TabIndex        =   2
            Top             =   1050
            Width           =   6165
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   390
            Width           =   975
            _Version        =   196608
            _ExtentX        =   1720
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
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
            Text            =   "2007"
            MaxValue        =   "9999"
            MinValue        =   "2007"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2685
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   390
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Mes de Informe:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Archivo a generar:"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   720
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   1244
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
            Height          =   585
            Left            =   630
            TabIndex        =   11
            Top             =   30
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Procesos Mivivienda - Informe de Pagos"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            Picture         =   "OpeTra_frm_057.frx":0890
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_MViPag_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_PerMes_Click()
   Call gs_SetFocus(ipp_PerAno)
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PerMes_Click
   End If
End Sub

Private Sub cmd_Proces_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de ejecutar el proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11

   'Buscando Pagos de Cliente y Procesar en Cronograma de Mivivienda
   Call fs_Buscar_Pagos_1
   
   'Generando Archivo de Texto con información de Pagos
   Call fs_Reporta_Pagos
   
   Screen.MousePointer = 0
   MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub drv_Listas_Change()
   dir_Carpet.Path = drv_Listas.Drive
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Value = Year(date)
   dir_Carpet.Path = "C:\"
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(dir_Carpet)
   End If
End Sub

Private Sub fs_Buscar_Pagos()
   Dim r_str_FecIni     As String
   Dim r_str_FecFin     As String
   Dim r_str_FecVct     As String
   Dim r_str_FecVc1     As String
   Dim r_int_PerMes     As Integer
   Dim r_int_PerAno     As Integer
   Dim r_str_NumOpe     As String
   Dim r_str_FecPag     As String
   Dim r_dbl_ImpPag     As Double
   Dim r_dbl_CapPag     As Double
   Dim r_dbl_IntPag     As Double
   Dim r_dbl_DesPag     As Double
   Dim r_dbl_VivPag     As Double
   Dim r_dbl_OtrPag     As Double
   Dim r_dbl_ICoPag     As Double
   Dim r_dbl_IMoPag     As Double
   Dim r_dbl_GCoPag     As Double
   Dim r_dbl_OtGPag     As Double
   
   'Se informa con Mes de Informe
   
   r_int_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex) - 1
   r_int_PerAno = ipp_PerAno.Value
   
   If r_int_PerMes = 0 Then
      r_int_PerMes = 12
      r_int_PerAno = r_int_PerAno - 1
   End If
   
   r_str_FecIni = Format(r_int_PerAno, "0000") & Format(r_int_PerMes, "00") & "16"
   r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "15"
   r_str_FecVct = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "01"
   r_str_FecVc1 = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"

   'Pagos antes del Vencimiento
   'Buscar Cuotas con Fecha de Pago antes del 15 del Mes Anterior
   'con Vencimiento entre el 01 del Mes Proceso al Ultimo Dia del Mes Proceso
   'Ejemplo:            Período Setiembre 2007
   'Fecha Pago          Inicial 15-08-2007 (Ini)
   'Fecha Pago          Final   16-09-2007 (Fin)
   'Fecha Vencimiento   Inicial 01-09-2007 (Vc1)
   'Fecha Vencimiento   Final   01-10-2007 (Vct)
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "SUBSTR(HIPCUO_NUMOPE,1,3) = '001' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1  AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECPAG < " & r_str_FecIni & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECPAG > 0 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & r_str_FecVc1 & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT < " & r_str_FecVct & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMOPE ASC, HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumOpe = g_rst_Princi!HIPCUO_NUMOPE
         r_str_FecVct = CStr(g_rst_Princi!HIPCUO_FECVCT)
         r_str_FecPag = CStr(g_rst_Princi!HIPCUO_FECPAG)
         
         r_dbl_ImpPag = g_rst_Princi!HIPCUO_IMPPAG
         r_dbl_CapPag = g_rst_Princi!HIPCUO_CAPPAG
         r_dbl_IntPag = g_rst_Princi!HIPCUO_INTPAG
         r_dbl_DesPag = g_rst_Princi!HIPCUO_DESPAG
         r_dbl_VivPag = g_rst_Princi!HIPCUO_VIVPAG
         r_dbl_OtrPag = g_rst_Princi!HIPCUO_OTRPAG
         r_dbl_ICoPag = g_rst_Princi!HIPCUO_ICOPAG
         r_dbl_IMoPag = g_rst_Princi!HIPCUO_IMOPAG
         r_dbl_GCoPag = g_rst_Princi!HIPCUO_GCOPAG
         r_dbl_OtGPag = g_rst_Princi!HIPCUO_OTGPAG
         
         'Actualizar en Cronograma de Pagos con Mivivienda
         Call fs_Pago_CrcPbp(r_str_NumOpe, r_str_FecVct, r_str_FecPag, r_dbl_ImpPag, r_dbl_CapPag, r_dbl_IntPag, r_dbl_DesPag, r_dbl_VivPag, r_dbl_OtrPag, r_dbl_ICoPag, r_dbl_IMoPag, r_dbl_GCoPag, r_dbl_OtGPag)
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Buscar Cuotas con Fecha de Pago entre el 15 del Mes Anterior y el 16 del Mes de Proceso
   'con Vencimiento entre el 01 del Mes Proceso al Ultimo Dia del Mes Proceso
   'Ejemplo:            Período Setiembre 2007
   'Fecha Pago          Inicial 15-08-2007 (Ini)
   'Fecha Pago          Final   16-09-2007 (Fin)
   'Fecha Vencimiento   Inicial 01-09-2007 (Vc1)
   'Fecha Vencimiento   Final   01-10-2007 (Vct)
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "SUBSTR(HIPCUO_NUMOPE,1,3) = '001' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1  AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECPAG >= " & r_str_FecIni & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECPAG <= " & r_str_FecFin & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & r_str_FecVc1 & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT < " & r_str_FecVct & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMOPE ASC, HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumOpe = g_rst_Princi!HIPCUO_NUMOPE
         r_str_FecVct = CStr(g_rst_Princi!HIPCUO_FECVCT)
         r_str_FecPag = CStr(g_rst_Princi!HIPCUO_FECPAG)
         
         r_dbl_ImpPag = g_rst_Princi!HIPCUO_IMPPAG
         r_dbl_CapPag = g_rst_Princi!HIPCUO_CAPPAG
         r_dbl_IntPag = g_rst_Princi!HIPCUO_INTPAG
         r_dbl_DesPag = g_rst_Princi!HIPCUO_DESPAG
         r_dbl_VivPag = g_rst_Princi!HIPCUO_VIVPAG
         r_dbl_OtrPag = g_rst_Princi!HIPCUO_OTRPAG
         r_dbl_ICoPag = g_rst_Princi!HIPCUO_ICOPAG
         r_dbl_IMoPag = g_rst_Princi!HIPCUO_IMOPAG
         r_dbl_GCoPag = g_rst_Princi!HIPCUO_GCOPAG
         r_dbl_OtGPag = g_rst_Princi!HIPCUO_OTGPAG
         
         'Actualizar en Cronograma de Pagos con Mivivienda
         Call fs_Pago_CrcPbp(r_str_NumOpe, r_str_FecVct, r_str_FecPag, r_dbl_ImpPag, r_dbl_CapPag, r_dbl_IntPag, r_dbl_DesPag, r_dbl_VivPag, r_dbl_OtrPag, r_dbl_ICoPag, r_dbl_IMoPag, r_dbl_GCoPag, r_dbl_OtGPag)
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'Buscar Cuotas con Período de Gracia
   'Con Fecha de Vencimiento entre el 01 y el 30 del Mes de Proceso
   r_str_FecIni = Format(r_int_PerAno, "0000") & Format(r_int_PerMes + 1, "00") & "01"
   r_str_FecFin = Format(r_int_PerAno, "0000") & Format(r_int_PerMes + 1, "00") & Format(ff_Ultimo_Dia_Mes(r_int_PerMes + 1, r_int_PerAno), "00")
   
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "SUBSTR(HIPCUO_NUMOPE,1,3) = '001' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 3  AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & r_str_FecIni & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & r_str_FecFin & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_CAPITA = 0 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumOpe = g_rst_Princi!HIPCUO_NUMOPE
         r_str_FecVct = CStr(g_rst_Princi!HIPCUO_FECVCT)
         r_str_FecPag = Format(date, "yyyymmdd")
         
         Call fs_Pago_CrcPbp(r_str_NumOpe, r_str_FecVct, r_str_FecVct, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Pago_CrcPbp(ByVal p_NumOpe As String, ByVal p_FecVct As String, ByVal p_FecPag As String, ByVal p_ImpPag As Double, ByVal p_CapPag As Double, ByVal p_IntPag As Double, ByVal p_DesPag As Double, ByVal p_VivPag As Double, ByVal p_OtrPag As Double, ByVal p_ICoPag As Double, ByVal p_IMoPag As Double, ByVal p_GCoPag As Double, ByVal p_OtGPag As Double)
   'Actualizar en Cronograma de Pagos con Mivivienda
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPCUO_CRCPBP ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumOpe & "', "
      g_str_Parame = g_str_Parame & p_FecVct & ", "
      g_str_Parame = g_str_Parame & p_FecPag & ", "
      g_str_Parame = g_str_Parame & CStr(p_ImpPag) & ", "
      g_str_Parame = g_str_Parame & CStr(p_CapPag) & ", "
      g_str_Parame = g_str_Parame & CStr(p_IntPag) & ", "
      g_str_Parame = g_str_Parame & CStr(p_DesPag) & ", "
      g_str_Parame = g_str_Parame & CStr(p_VivPag) & ", "
      g_str_Parame = g_str_Parame & CStr(p_OtrPag) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ICoPag) & ", "
      g_str_Parame = g_str_Parame & CStr(p_IMoPag) & ", "
      g_str_Parame = g_str_Parame & CStr(p_GCoPag) & ", "
      g_str_Parame = g_str_Parame & CStr(p_OtGPag) & ", "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                              'Código Sucursal
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_HIPCUO_CRCPBP. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
End Sub

Private Sub fs_Reporta_Pagos()
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_str_FecVct     As String
Dim r_str_FecVc1     As String
Dim r_int_PerMes     As Integer
Dim r_int_PerAno     As Integer
Dim r_str_NumOpe     As String
Dim r_str_FecPag     As String
Dim r_str_OpeMVi     As String
Dim r_dbl_SalCon     As Double
Dim r_dbl_SalIns     As Double
Dim r_int_NumFil     As Integer
Dim r_int_NumCuo     As Integer
Dim r_dbl_SalCap     As Double
Dim r_int_PerGra     As Integer
Dim r_dbl_MtoGra     As Double
Dim r_dbl_MtoPre     As Double

   'Buscar Cuotas Pagadas por el Cliente
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE (SUBSTR(HIPCUO_NUMOPE,1,3) = '001' OR SUBSTR(HIPCUO_NUMOPE,1,3) = '003') "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 3 "
   g_str_Parame = g_str_Parame & "   AND SEGFECACT = " & Format(date, "yyyymmdd")
   g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMOPE ASC, HIPCUO_NUMCUO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_NumFil = FreeFile
      Open dir_Carpet.Path & "\P" & Format(date, "yyyymmdd") & ".064" For Output As r_int_NumFil
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_str_NumOpe = g_rst_Princi!HIPCUO_NUMOPE
         r_int_NumCuo = g_rst_Princi!HIPCUO_NUMCUO
         r_str_FecPag = CStr(g_rst_Princi!HIPCUO_FECPAG)
         r_dbl_SalCap = g_rst_Princi!HIPCUO_SALCAP
         r_str_OpeMVi = ""
         r_dbl_SalCon = 0
         r_int_PerGra = 0
         r_dbl_MtoGra = 0
         
         'Buscar Nro de Operación Mivivienda y Saldo del Tramo Concesional
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT * "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
         g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & r_str_NumOpe & "' "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            r_str_OpeMVi = Trim(g_rst_Genera!HIPMAE_OPEMVI)
            r_dbl_SalCon = g_rst_Genera!HIPMAE_SALCON
            r_int_PerGra = g_rst_Genera!HIPMAE_PERGRA
            r_dbl_MtoPre = g_rst_Genera!HIPMAE_MTOMVI
            r_dbl_MtoGra = g_rst_Genera!HIPMAE_IMPCON + g_rst_Genera!HIPMAE_IMPNCO
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         r_dbl_SalIns = r_dbl_SalCap + r_dbl_SalCon
         
         If r_int_PerGra > 0 Then
            If g_rst_Princi!HIPCUO_NUMCUO < r_int_PerGra Then
               r_dbl_SalIns = r_dbl_MtoPre
            ElseIf g_rst_Princi!HIPCUO_NUMCUO = r_int_PerGra Then
               r_dbl_SalIns = r_dbl_MtoGra
            End If
         End If
         
         'Escribiendo en Archivo de Texto
         Print #r_int_NumFil, Mid(r_str_OpeMVi & Space(16), 1, 16) & " " & _
                              Format(r_int_NumCuo, "000") & " " & _
                              r_str_FecPag & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(r_dbl_SalIns, "########0.00"), 2))) & gf_ComaDecimal(Format(r_dbl_SalIns, "########0.00"), 2)
         
         g_rst_Princi.MoveNext
      Loop
      
      'Cerrando Archivo Cabecera
      Close #r_int_NumFil
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_Pagos_1()
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_str_FecVcI     As String
Dim r_str_FecVcF     As String
Dim r_str_FecVct     As String
Dim r_int_MesIni     As Integer
Dim r_int_AnoIni     As Integer
Dim r_int_MesFin     As Integer
Dim r_int_AnoFin     As Integer
Dim r_str_NumOpe     As String
Dim r_str_FecPag     As String
Dim r_dbl_ImpPag     As Double
Dim r_dbl_CapPag     As Double
Dim r_dbl_IntPag     As Double
Dim r_dbl_DesPag     As Double
Dim r_dbl_VivPag     As Double
Dim r_dbl_OtrPag     As Double
Dim r_dbl_ICoPag     As Double
Dim r_dbl_IMoPag     As Double
Dim r_dbl_GCoPag     As Double
Dim r_dbl_OtGPag     As Double
   
   'Se informa con Mes de Informe
   If cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 1 Then
      r_int_MesIni = 12
      r_int_AnoIni = ipp_PerAno.Value - 1
   Else
      r_int_MesIni = cmb_PerMes.ItemData(cmb_PerMes.ListIndex) - 1
      r_int_AnoIni = ipp_PerAno.Value
   End If
   
   If cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 12 Then
      r_int_MesFin = 1
      r_int_AnoFin = ipp_PerAno.Value + 1
   Else
      r_int_MesFin = cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1
      r_int_AnoFin = ipp_PerAno.Value
   End If
   
   'Para Rango de Pagos
   r_str_FecIni = Format(r_int_AnoIni, "0000") & Format(r_int_MesIni, "00") & "16"
   r_str_FecFin = Format(r_int_AnoFin, "0000") & Format(r_int_MesFin, "00") & "15"
   
   'Para Rangos de Vencimientos
   r_str_FecVcI = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
   r_str_FecVcF = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")

   'Pagos antes del Vencimiento
   'Buscar Cuotas con Fecha de Pago antes del 15 del Mes Anterior
   'con Vencimiento entre el 01 del Mes Proceso al Ultimo Dia del Mes Proceso
   'Ejemplo:            Período Setiembre 2007
   'Fecha Pago          Inicial 16-08-2007 (Ini)
   'Fecha Vencimiento   Inicial 01-09-2007 (Vc1)
   'Fecha Vencimiento   Final   30-09-2007 (Vct)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE (SUBSTR(HIPCUO_NUMOPE,1,3) = '001' OR SUBSTR(HIPCUO_NUMOPE,1,3) = '003') "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECPAG < " & r_str_FecFin & " "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECPAG > 0 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT >= " & r_str_FecVcI & " "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & r_str_FecVcF & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMOPE ASC, HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_str_NumOpe = g_rst_Princi!HIPCUO_NUMOPE
         r_str_FecVct = CStr(g_rst_Princi!HIPCUO_FECVCT)
         r_str_FecPag = CStr(g_rst_Princi!HIPCUO_FECPAG)
         r_dbl_ImpPag = g_rst_Princi!HIPCUO_IMPPAG
         r_dbl_CapPag = g_rst_Princi!HIPCUO_CAPPAG
         r_dbl_IntPag = g_rst_Princi!HIPCUO_INTPAG
         r_dbl_DesPag = g_rst_Princi!HIPCUO_DESPAG
         r_dbl_VivPag = g_rst_Princi!HIPCUO_VIVPAG
         r_dbl_OtrPag = g_rst_Princi!HIPCUO_OTRPAG
         r_dbl_ICoPag = g_rst_Princi!HIPCUO_ICOPAG
         r_dbl_IMoPag = g_rst_Princi!HIPCUO_IMOPAG
         r_dbl_GCoPag = g_rst_Princi!HIPCUO_GCOPAG
         r_dbl_OtGPag = g_rst_Princi!HIPCUO_OTGPAG
         
         'Actualizar en Cronograma de Pagos con Mivivienda
         Call fs_Pago_CrcPbp(r_str_NumOpe, r_str_FecVct, r_str_FecPag, r_dbl_ImpPag, r_dbl_CapPag, r_dbl_IntPag, r_dbl_DesPag, r_dbl_VivPag, r_dbl_OtrPag, r_dbl_ICoPag, r_dbl_IMoPag, r_dbl_GCoPag, r_dbl_OtGPag)
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Buscar Cuotas con Fecha de Pago (Mes Proceso Anterior al Mes Actual)
   'con Vencimiento entre el 01 del Mes Proceso al Ultimo Dia del Mes Proceso
   'Ejemplo:            Período Setiembre 2007
   'Fecha Pago          Inicial 16-08-2007 (Ini)
   'Fecha Pago          Final   15-10-2007 (Fin)
   'Fecha Vencimiento   Inicial 01-09-2007 (Vc1)
   'Fecha Vencimiento   Final   30-09-2007 (Vct)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE (SUBSTR(HIPCUO_NUMOPE,1,3) = '001' OR SUBSTR(HIPCUO_NUMOPE,1,3) = '003') "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECPAG >= " & r_str_FecIni & " "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECPAG <= " & r_str_FecFin & " "
   'g_str_Parame = g_str_Parame & "  AND HIPCUO_FECVCT >= " & r_str_FecVcI & " "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & r_str_FecVcF & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMOPE ASC, HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_str_NumOpe = g_rst_Princi!HIPCUO_NUMOPE
         r_str_FecVct = CStr(g_rst_Princi!HIPCUO_FECVCT)
         r_str_FecPag = CStr(g_rst_Princi!HIPCUO_FECPAG)
         r_dbl_ImpPag = g_rst_Princi!HIPCUO_IMPPAG
         r_dbl_CapPag = g_rst_Princi!HIPCUO_CAPPAG
         r_dbl_IntPag = g_rst_Princi!HIPCUO_INTPAG
         r_dbl_DesPag = g_rst_Princi!HIPCUO_DESPAG
         r_dbl_VivPag = g_rst_Princi!HIPCUO_VIVPAG
         r_dbl_OtrPag = g_rst_Princi!HIPCUO_OTRPAG
         r_dbl_ICoPag = g_rst_Princi!HIPCUO_ICOPAG
         r_dbl_IMoPag = g_rst_Princi!HIPCUO_IMOPAG
         r_dbl_GCoPag = g_rst_Princi!HIPCUO_GCOPAG
         r_dbl_OtGPag = g_rst_Princi!HIPCUO_OTGPAG
         
         'Actualizar en Cronograma de Pagos con Mivivienda
         Call fs_Pago_CrcPbp(r_str_NumOpe, r_str_FecVct, r_str_FecPag, r_dbl_ImpPag, r_dbl_CapPag, r_dbl_IntPag, r_dbl_DesPag, r_dbl_VivPag, r_dbl_OtrPag, r_dbl_ICoPag, r_dbl_IMoPag, r_dbl_GCoPag, r_dbl_OtGPag)
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Buscar Cuotas con Período de Gracia
   'Con Fecha de Vencimiento entre el 01 y el 30 del Mes de Proceso
   'r_str_FecIni = Format(r_int_PerAno, "0000") & Format(r_int_PerMes + 1, "00") & "01"
   'r_str_FecFin = Format(r_int_PerAno, "0000") & Format(r_int_PerMes + 1, "00") & Format(ff_Ultimo_Dia_Mes(r_int_PerMes + 1, r_int_PerAno), "00")
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE (SUBSTR(HIPCUO_NUMOPE,1,3) = '001' OR SUBSTR(HIPCUO_NUMOPE,1,3) = '003') "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 3 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT >= " & r_str_FecIni & " "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & r_str_FecFin & " "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_CAPITA = 0 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_str_NumOpe = g_rst_Princi!HIPCUO_NUMOPE
         r_str_FecVct = CStr(g_rst_Princi!HIPCUO_FECVCT)
         
         Call fs_Pago_CrcPbp(r_str_NumOpe, r_str_FecVct, r_str_FecVct, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
