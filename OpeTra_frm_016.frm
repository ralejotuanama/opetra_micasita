VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm_Pro_VisSbs_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2415
   ClientLeft      =   2700
   ClientTop       =   3795
   ClientWidth     =   9930
   Icon            =   "OpeTra_frm_016.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2415
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9915
      _Version        =   65536
      _ExtentX        =   17489
      _ExtentY        =   4260
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
         Height          =   795
         Left            =   30
         TabIndex        =   7
         Top             =   750
         Width           =   9825
         _Version        =   65536
         _ExtentX        =   17330
         _ExtentY        =   1402
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1470
            TabIndex        =   1
            Top             =   420
            Width           =   1005
            _Version        =   196608
            _ExtentX        =   1773
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
            ButtonStyle     =   1
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
            Text            =   "0"
            MaxValue        =   "9999"
            MinValue        =   "2005"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin VB.Label Label6 
            Caption         =   "Ingrese Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   9
            Top             =   420
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Seleccione Mes:"
            Height          =   315
            Left            =   90
            TabIndex        =   8
            Top             =   90
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   9825
         _Version        =   65536
         _ExtentX        =   17330
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
            TabIndex        =   6
            Top             =   60
            Width           =   4965
            _Version        =   65536
            _ExtentX        =   8758
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Vista SBS - Anexo I"
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
            Left            =   60
            Picture         =   "OpeTra_frm_016.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   10
         Top             =   1590
         Width           =   9825
         _Version        =   65536
         _ExtentX        =   17330
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   9120
            Picture         =   "OpeTra_frm_016.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   8430
            Picture         =   "OpeTra_frm_016.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   30
            Width           =   675
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_VisSbs_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Imprim_Click()
   Dim r_int_NumFil     As Integer
   Dim r_int_Contad     As Integer
   
   'On Error GoTo cmd_ArcTxt_Error
   
   If MsgBox("¿Está seguro de Guardar el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call fs_Genera
   
   dlg_Guarda.Filter = "Texto (*.txt)|*.txt"
   dlg_Guarda.ShowSave
   
   
   'Crear Archivo
   r_int_NumFil = FreeFile
   Open dlg_Guarda.FileName For Output As r_int_NumFil
                
   For r_int_Contad = 1 To UBound(g_arr_Imprim)
      If g_arr_Imprim(r_int_Contad).Imprim_ConLen = "SP" Then
         Print #r_int_NumFil, Chr(12)
      Else
         If Len(Trim(g_arr_Imprim(r_int_Contad).Imprim_ConLen)) > 0 Then
            Print #r_int_NumFil, g_arr_Imprim(r_int_Contad).Imprim_ConLen
         Else
            Print #r_int_NumFil, ""
         End If
      End If
   Next r_int_Contad
                
   'Cerrando Archivo
   Close #r_int_NumFil
   
cmd_ArcTxt_Error:
   Exit Sub
   
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Value = Year(Date)
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Genera()
   Dim r_str_CodPrd     As String
   Dim r_str_NumOpe     As String
   Dim r_str_NumSol     As String
   Dim r_str_TipDoc     As String
   Dim r_str_NumDoc     As String
   Dim r_str_CodCli     As String
   Dim r_str_EjeVta     As String
   Dim r_str_DesSuc     As String
   Dim r_str_CodSbs     As String
   Dim r_str_TipCre     As String
   Dim r_str_IndCre     As String
   Dim r_str_CodCiu     As String
   Dim r_str_TipMon     As String
   Dim r_str_MtoPre     As String
   Dim r_str_FecDes     As String
   Dim r_str_NumCuo     As String
   Dim r_str_PerCuo     As String
   Dim r_str_NumRen     As String
   Dim r_str_SalCap     As String
   Dim r_str_PrxVct     As String
   Dim r_str_DiaAtr     As String
   Dim r_str_EstCre     As String
   Dim r_str_CapVen     As String
   Dim r_str_CalCre     As String
   Dim r_str_IntDev     As String
   Dim r_str_IntPer     As String
   Dim r_str_IntSus     As String
   Dim r_str_TotGar     As String
   Dim r_str_TipGar     As String
   Dim r_str_NomPer     As String
   Dim r_str_AtrGar     As String
   Dim r_str_NumGar     As String
   Dim r_str_FecTas     As String
   Dim r_str_PrvReq     As String
   Dim r_str_PrvCon     As String
   Dim r_int_NumGar     As Integer
   Dim r_dbl_TotGar     As Double
   Dim r_dbl_CapVen     As Double
   Dim r_dbl_IntDev     As Double
   Dim r_dbl_IntSus     As Double
   Dim r_dbl_PrvReq     As Double
   Dim r_dbl_PrvCon     As Double
   Dim r_int_UltDia     As Integer

   ReDim g_arr_Imprim(0)

   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 OR "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 3  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   Call gs_LinImp("")
   Call gs_LinImp("")
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
      r_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
      r_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
      
      r_str_CodSbs = Space(30)
      
      'Obtener Información de Maestro de Productos
      g_str_Parame = "SELECT * FROM CRE_PRODUC WHERE "
      g_str_Parame = g_str_Parame & "PRODUC_CODIGO = '" & r_str_CodPrd & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         r_str_TipCre = Mid(Mid(moddat_gf_Consulta_Pardes("055", CStr(g_rst_GenAux!PRODUC_CODCLA)), 5) & Space(30), 1, 30)
         r_str_IndCre = Mid(Mid(moddat_gf_Consulta_Pardes("056", CStr(g_rst_GenAux!PRODUC_TIPCRE)), 5) & Space(30), 1, 30)
      End If
      
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      
      
      'Obtener Información de Maestro de Clientes
      g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(g_rst_Princi!HIPMAE_TDOCLI) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & g_rst_Princi!HIPMAE_NDOCLI & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         r_str_CodSbs = Mid(Trim(g_rst_GenAux!DatGen_CodSbs & "") & Space(30), 1, 30)
      End If
      
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      
      'Obtener Información de Tasacion
      g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
      g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & r_str_NumSol & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         r_str_FecTas = gf_FormatoFecha(CStr(g_rst_GenAux!EVATAS_FECEVA))
         r_str_NomPer = Mid(Trim(g_rst_GenAux!EVATAS_NOMPER) & Space(60), 1, 60)
      End If
      
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      
      
      r_str_TotGar = gf_FormatoNumero(0, 15, 2)
      r_str_TipGar = Space(30)
      r_str_AtrGar = Space(30)
      r_str_NumGar = gf_FormatoNumEnt(0, 5)
      r_int_NumGar = 0
      r_dbl_TotGar = 0
      
      'Obtener Información de Garantías
      g_str_Parame = "SELECT * FROM CRE_HIPGAR WHERE "
      g_str_Parame = g_str_Parame & "HIPGAR_NUMOPE = '" & r_str_NumOpe & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         g_rst_GenAux.MoveFirst
         
         
         Do While Not g_rst_GenAux.EOF
            If Len(Trim(r_str_TipGar)) = 0 Then
               r_str_TipGar = Mid(moddat_gf_Consulta_TipGar(Right(g_rst_GenAux!HIPGAR_TIPGAR, 2)) & Space(30), 1, 30)
               r_str_AtrGar = Mid(moddat_gf_Consulta_PreGar(Right(g_rst_GenAux!HIPGAR_TIPGAR, 2)) & Space(30), 1, 30)
            End If
         
            r_int_NumGar = r_int_NumGar + 1
            r_dbl_TotGar = r_dbl_TotGar + CDbl(g_rst_GenAux!HIPGAR_MTOHIP)
            g_rst_GenAux.MoveNext
         Loop
      End If
      
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      
      r_str_TotGar = gf_FormatoNumero(r_dbl_TotGar, 15, 2)
      r_str_NumGar = gf_FormatoNumEnt(r_int_NumGar, 5)
      
      
      
      'Obtener Calificación
      r_dbl_PrvReq = 0
      r_dbl_PrvCon = 0
      r_str_CalCre = Space(30)
      
      g_str_Parame = "SELECT * FROM AUDIT_PROVISIONES WHERE "
      g_str_Parame = g_str_Parame & "CREDITO = '" & r_str_NumOpe & "' AND "
      g_str_Parame = g_str_Parame & "ANO = " & CStr(ipp_PerAno.Value) & " AND "
      g_str_Parame = g_str_Parame & "MES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         r_str_CalCre = Mid(Mid(moddat_gf_Consulta_Pardes("058", CStr(CInt(g_rst_GenAux!CALIFICACION) + 1)), 5) & Space(30), 1, 30)
         r_dbl_PrvReq = g_rst_GenAux!SALDO_CAPITAL * g_rst_GenAux!TASA_PROVISION
         r_dbl_PrvCon = g_rst_GenAux!SALDO_CAPITAL * g_rst_GenAux!TASA_PROVISION
      End If
      
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      
      
      'Obtener Capital Vencido
      r_dbl_CapVen = 0
      
      g_str_Parame = "SELECT * FROM CREDITO_CIERRE_FINMES WHERE "
      g_str_Parame = g_str_Parame & "CREDITO = '" & r_str_NumOpe & "' AND "
      g_str_Parame = g_str_Parame & "ANO = " & CStr(ipp_PerAno.Value) & " AND "
      g_str_Parame = g_str_Parame & "MES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         r_dbl_CapVen = g_rst_GenAux!CAPITAL_VENCIDO
      End If
      
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      
      'Obtener Interes Devengado / Suspenso
      r_dbl_IntDev = 0
      r_dbl_IntSus = 0
      r_int_UltDia = ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Value)
      
      g_str_Parame = "SELECT * FROM CRED_DEVENGADO_INTERES WHERE "
      g_str_Parame = g_str_Parame & "CREDITO = '" & r_str_NumOpe & "' AND "
      g_str_Parame = g_str_Parame & "FECHA = to_date(" & Format(CDate(Format(r_int_UltDia, "00") & "/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Value, "0000")), "yyyymmdd") & ", 'yyyy/mm/dd')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         If g_rst_Princi!HIPMAE_SITCRE = 4 Then
            r_dbl_IntSus = g_rst_GenAux!INTERES
         Else
            r_dbl_IntDev = g_rst_GenAux!INTERES
         End If
      End If
      
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
            
      
      'Creando Línea de Archivo
      r_str_TipDoc = Mid(moddat_gf_Consulta_Pardes("203", CStr(g_rst_Princi!HIPMAE_TDOCLI)) & Space(30), 1, 30)
      r_str_NumDoc = Left(Trim(g_rst_Princi!HIPMAE_NDOCLI) & Space(12), 12)
      r_str_CodCli = CStr(g_rst_Princi!HIPMAE_TDOCLI) & r_str_NumDoc
      r_str_DesSuc = "PRINCIPAL"
      r_str_EjeVta = Mid(moddat_gf_Buscar_NomEje(g_rst_Princi!HIPMAE_EJEVTA) & Space(30), 1, 30)
      r_str_CodCiu = Format(g_rst_Princi!HIPMAE_CODCIU, "0000")
      r_str_MtoPre = gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 15, 2)
      r_str_TipMon = Mid(moddat_gf_Consulta_Pardes("204", CStr(g_rst_Princi!HIPMAE_MONEDA)) & Space(30), 1, 30)
      r_str_FecDes = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      r_str_NumCuo = Format(g_rst_Princi!HIPMAE_NUMCUO, "000")
      r_str_PerCuo = "MENSUAL"
      r_str_NumRen = "0"
      r_str_SalCap = gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 15, 2)
      r_str_PrxVct = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_PRXVCT))
      r_str_DiaAtr = Format(g_rst_Princi!HIPMAE_DIAMOR, "0000")
      r_str_EstCre = Mid(Mid(moddat_gf_Consulta_Pardes("057", CStr(g_rst_Princi!HIPMAE_SITCRE)), 5) & Space(30), 1, 30)
      
      r_str_IntPer = gf_FormatoNumero(g_rst_Princi!HIPMAE_PAGINT, 15, 2)
      
      r_str_CapVen = gf_FormatoNumero(r_dbl_CapVen, 15, 2)
      r_str_IntDev = gf_FormatoNumero(r_dbl_IntDev, 15, 2)
      r_str_IntSus = gf_FormatoNumero(r_dbl_IntSus, 15, 2)
      r_str_PrvReq = gf_FormatoNumero(r_dbl_PrvReq, 15, 2)
      r_str_PrvCon = gf_FormatoNumero(r_dbl_PrvCon, 15, 2)

      Call gs_LinImp(r_str_TipDoc & " " & r_str_NumDoc & " " & r_str_CodCli & " " & r_str_CodSbs & " " & r_str_DesSuc & " " & r_str_NumOpe & " " & r_str_TipCre & " " & r_str_IndCre & " " & r_str_MtoPre & " " & r_str_TipMon & " " & r_str_FecDes & " " & r_str_NumCuo & " " & r_str_PerCuo & " " & r_str_NumRen & " " & r_str_SalCap & " " & r_str_PrxVct & " " & r_str_DiaAtr & " " & r_str_EstCre & " --- " & r_str_CapVen & " " & r_str_CalCre & " " & r_str_IntDev & " " & r_str_IntPer & " " & r_str_IntSus & " " & r_str_TotGar & " " & r_str_TipGar & " " & r_str_NomPer & " " & r_str_AtrGar & " " & r_str_NumGar & " " & r_str_FecTas & " " & r_str_PrvReq & " " & r_str_PrvCon)
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
