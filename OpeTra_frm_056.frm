VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Caj_CarArc_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   4440
   ClientTop       =   3870
   ClientWidth     =   7890
   Icon            =   "OpeTra_frm_056.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2295
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7905
      _Version        =   65536
      _ExtentX        =   13944
      _ExtentY        =   4048
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
         Height          =   645
         Left            =   30
         TabIndex        =   11
         Top             =   780
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
            Left            =   7200
            Picture         =   "OpeTra_frm_056.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Import 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_056.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   6
         Top             =   1470
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
            Left            =   7440
            TabIndex        =   2
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txt_NomArc 
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   5835
         End
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   6195
         End
         Begin VB.Label Label1 
            Caption         =   "Archivo a cargar:"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   390
            Width           =   1365
         End
         Begin VB.Label Label3 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   30
         TabIndex        =   8
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
            Height          =   315
            Left            =   630
            TabIndex        =   9
            Top             =   30
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Operaciones Financieras"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   315
            Left            =   630
            TabIndex        =   12
            Top             =   330
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7064
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Carga de Archivo de Recaudación"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   7200
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
            Left            =   6570
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   5550
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   6090
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_056.frx":0890
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Caj_CarArc_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_CodBan()      As moddat_tpo_Genera
Dim l_arr_GasCie()      As moddat_tpo_Genera
Dim l_dbl_PorITF        As Double
Dim l_str_CodRub        As String
Dim l_str_CodEmp        As String
Dim l_str_CodSer        As String
Dim l_str_CodBan        As String

Private Sub cmd_Import_Click()
   If cmb_CodBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodBan)
      Exit Sub
   End If
   If Len(Trim(txt_NomArc.Text)) = 0 Then
      MsgBox "Debe ingresar la ubicación y nombre del archivo a importar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de ejecutar el proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   l_str_CodBan = l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo
   
   Screen.MousePointer = 11
   Select Case l_str_CodBan
      Case "000002": Call fs_Carga_BBVA(txt_NomArc.Text)
      Case "000004": Call fs_Carga_Interbank(txt_NomArc.Text)
   End Select
   Screen.MousePointer = 0
End Sub

Private Sub cmd_BusArc_Click()
   On Error GoTo cmd_BusArc_Error
   dlg_Guarda.Filter = "Todos los archivos (*.*)|*.*"
   dlg_Guarda.ShowOpen
   txt_NomArc.Text = UCase(dlg_Guarda.FileName)
   Exit Sub
   
cmd_BusArc_Error:
   txt_NomArc.Text = ""
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_CodBan)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
   txt_NomArc.Text = ""
End Sub

Private Sub cmb_CodBan_Click()
   Call gs_SetFocus(txt_NomArc)
End Sub

Private Sub cmb_CodBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodBan_Click
   End If
End Sub

Private Sub txt_NomArc_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_.\")
   Else
      Call gs_SetFocus(cmd_Import)
   End If
End Sub

Private Sub fs_Carga_Interbank(ByVal p_NomFil As String)
Dim r_int_NumFil     As Integer
Dim r_str_Cadena     As String
Dim r_str_CodBco     As String
Dim r_str_NumCta     As String
Dim r_str_FecRec     As String
Dim r_str_NumOpe     As String
Dim r_str_FecPag     As String
Dim r_str_OfiPag     As String
Dim r_str_ForPag     As String
Dim r_str_CanPag     As String
Dim r_str_NumDoc     As String
Dim r_int_TipDoc     As Integer
Dim r_int_TipMon     As Integer
Dim r_int_TipPag     As Integer
Dim r_int_NumCuo     As Integer
Dim r_dbl_ImpDep     As Double
Dim r_str_NumMov     As String
Dim r_str_CadErr     As String
Dim r_int_CntErr     As Integer
Dim r_int_FlgErr     As Integer
Dim r_str_OpeAux     As String
Dim r_str_DocIde     As String
   
   ' inicializa variables
   r_str_CodBco = "000004"          'Código de banco
   l_str_CodRub = "07"              'Código de rubro
   l_str_CodEmp = "245"             'Código de empresa
   l_str_CodSer = "01"              'Código de servicio
   r_str_NumCta = ""
   r_int_TipMon = 0
   r_str_FecRec = ""
   
   r_int_NumFil = FreeFile
   Open p_NomFil For Input As r_int_NumFil
   
   r_int_CntErr = 0
   r_int_FlgErr = 0
   r_str_OpeAux = ""
   ReDim l_arr_GasCie(0)
   
   Do While Not EOF(r_int_NumFil)
      Line Input #r_int_NumFil, r_str_Cadena
      DoEvents
      
      If Left(r_str_Cadena, 2) = l_str_CodRub And Mid(r_str_Cadena, 3, 3) = l_str_CodEmp And Mid(r_str_Cadena, 6, 2) = l_str_CodSer Then    'Rubro, Código de la Empresa y Código de Servicio
         
         r_str_FecRec = Mid(r_str_Cadena, 124, 8)                                                                    'Fecha de Emisión
         Select Case Mid(r_str_Cadena, 8, 2)                                                                         'Tipo de Moneda
            Case "10": r_int_TipMon = 2
            Case "01": r_int_TipMon = 1
         End Select
         
         r_str_DocIde = Format(Trim(Mid(r_str_Cadena, 10, 20)), "#############")
         r_int_TipDoc = CInt(Mid(r_str_DocIde, 1, 1))                                                                'Tipo de Documento
         r_str_NumDoc = Trim(Mid(r_str_DocIde, 2, 8))                                                                'Número de Documento
         r_int_TipPag = IIf(CInt(Mid(r_str_Cadena, 37, 1)) = 1, 2, 1)                                                'Tipo de Pago
         r_str_NumOpe = ff_BuscarNumero(r_int_TipDoc, r_str_NumDoc, r_int_TipPag)                                    'Número de Operación / Solicitud
         r_int_NumCuo = ff_BuscarNumCuo(r_str_NumOpe, Mid(r_str_Cadena, 30, 4) & Mid(r_str_Cadena, 35, 2))           'Número de la Cuota
         r_str_FecPag = Mid(r_str_Cadena, 83, 8)                                                                     'Fecha de Pago
         r_str_NumMov = Mid(r_str_Cadena, 140, 8)                                                                    'Número de Movimiento
         r_dbl_ImpDep = CDbl(Mid(r_str_Cadena, 97, 11) & "." & Mid(r_str_Cadena, 108, 2))                            'Importe Pagado
         r_str_OfiPag = "9999"                                                                                       'Oficina de Pago
         r_str_ForPag = ""
         r_str_CanPag = ""
         
         Select Case r_int_TipPag
            Case 1      'PAGO DE GASTOS DE CIERRE
               If Not ff_GasAdm(r_str_NumOpe, r_int_TipDoc, r_str_NumDoc, r_str_CodBco, r_str_FecPag, r_str_NumCta, r_str_NumMov, r_int_TipMon, r_dbl_ImpDep, r_str_FecRec, r_str_OfiPag, r_str_ForPag, r_str_CanPag, r_str_CadErr) Then
                  r_int_CntErr = r_int_CntErr + 1
                  Call fs_Graba_Report(gf_Formato_NumSol(r_str_NumOpe), r_int_TipDoc, r_str_NumDoc, r_str_FecPag, r_dbl_ImpDep, r_int_TipPag, 0, r_str_CadErr)
               End If
               
            Case 2      'PAGO DE CUOTAS
               If Not (r_int_FlgErr = 1 And r_str_OpeAux = r_str_NumOpe) Then
                  If Not ff_CuoHip(r_str_NumOpe, r_int_TipDoc, r_str_NumDoc, r_str_CodBco, r_str_FecPag, r_str_NumCta, r_str_NumMov, r_int_TipMon, r_dbl_ImpDep, r_str_FecRec, r_str_OfiPag, r_str_ForPag, r_str_CanPag, r_int_NumCuo, r_str_CadErr) Then
                     r_int_FlgErr = 1
                     r_str_OpeAux = r_str_NumOpe
                     r_int_CntErr = r_int_CntErr + 1
                     Call fs_Graba_Report(gf_Formato_NumOpe(r_str_NumOpe), r_int_TipDoc, r_str_NumDoc, r_str_FecPag, r_dbl_ImpDep, r_int_TipPag, r_int_NumCuo, r_str_CadErr)
                  Else
                     r_int_FlgErr = 0
                     r_str_OpeAux = ""
                  End If
               Else
                  Call fs_Graba_Report(gf_Formato_NumOpe(r_str_NumOpe), r_int_TipDoc, r_str_NumDoc, r_str_FecPag, r_dbl_ImpDep, r_int_TipPag, r_int_NumCuo, "Pago no procesado.")
               End If
         End Select
      End If
   Loop
   
   Close #r_int_NumFil
   Call fs_Correo_GasCie(cmb_CodBan.Text, moddat_gf_Consulta_ParDes("204", CStr(r_int_TipMon)))
   
   If r_int_CntErr > 0 Then
'      'Grabando en DAO
'      moddat_g_str_CadDAO = "SELECT * FROM RPT_CABGEN WHERE CABGEN_PRODUC = ' '"
'      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
'      moddat_g_rst_RecDAO.AddNew
'      moddat_g_rst_RecDAO("CABGEN_NOMBCO") = cmb_CodBan.Text
'      moddat_g_rst_RecDAO("CABGEN_NUMCTA") = r_str_NumCta
'      moddat_g_rst_RecDAO("CABGEN_FECPRO") = Format(date, "dd/mm/yyyy")
'      moddat_g_rst_RecDAO("CABGEN_FECREC") = gf_FormatoFecha(r_str_FecRec)
'      moddat_g_rst_RecDAO("CABGEN_MONEDA") = moddat_gf_Consulta_ParDes("204", CStr(r_int_TipMon))
'      moddat_g_rst_RecDAO.Update
'      DoEvents
'
'      moddat_g_rst_RecDAO.Close
'      DoEvents
'
'      'Generando Reporte
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_COBBCO_01.RPT"
'      crp_Imprim.Action = 1
      MsgBox "Error en la carga del archivo.", vbInformation, modgen_g_str_NomPlt
   Else
      MsgBox "El archivo fue cargado exitosamente.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Function ff_BuscarNumero(ByVal p_TdoCli As String, ByVal p_NdoCli As String, ByVal p_TipPag As Integer) As String
   ff_BuscarNumero = ""
   
   If p_TipPag = 2 Then
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_TDOCLI = " & p_TdoCli & " AND "
      g_str_Parame = g_str_Parame & "HIPMAE_NDOCLI = '" & p_NdoCli & "' AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
         Exit Function
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         ff_BuscarNumero = Trim(g_rst_Listas!HIPMAE_NUMOPE)
      End If
      
   ElseIf p_TipPag = 1 Then
   
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & p_TdoCli & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = " & p_NdoCli & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
         Exit Function
      End If
   
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         ff_BuscarNumero = Trim(g_rst_Listas!SOLMAE_NUMERO)
      End If
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function ff_BuscarNumCuo(ByVal p_NumOpe As String, ByVal p_FecVct As String) As Integer
   ff_BuscarNumCuo = 0
      
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = " & p_NumOpe & " AND "
   g_str_Parame = g_str_Parame & "SUBSTR(HIPCUO_FECVCT,1,6) = " & p_FecVct & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      ff_BuscarNumCuo = Trim(g_rst_Listas!HIPCUO_NUMCUO)
   End If

   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_Carga_BBVA(ByVal p_NomFil As String)
Dim r_int_NumFil     As Integer
Dim r_str_Cadena     As String
Dim r_str_CodBco     As String
Dim r_str_NumCta     As String
Dim r_str_FecRec     As String
Dim r_str_NumOpe     As String
Dim r_str_FecPag     As String
Dim r_str_OfiPag     As String
Dim r_str_ForPag     As String
Dim r_str_CanPag     As String
Dim r_str_NumDoc     As String
Dim r_int_TipDoc     As Integer
Dim r_int_TipMon     As Integer
Dim r_int_TipPag     As Integer
Dim r_int_NumCuo     As Integer
Dim r_dbl_ImpDep     As Double
Dim r_str_NumMov     As String
Dim r_str_CadErr     As String
Dim r_int_CntErr     As Integer
Dim r_int_FlgErr     As Integer
Dim r_str_OpeAux     As String
Dim r_str_DocIde     As String
Dim r_str_HorIni     As String
Dim r_str_HorFin     As String
Dim r_str_Situac     As String
Dim r_str_NomCli     As String
Dim r_int_NumReg     As Integer
Dim r_dbl_IMPTOT     As Double
Dim r_int_ConErr     As Integer
Dim r_int_SinErr     As Integer
Dim r_int_NumPro     As Integer
Dim r_int_ErrorUpd   As Integer
Dim r_int_NFila      As Integer

   r_str_CodBco = l_str_CodBan
   r_str_NumCta = ""
   r_int_TipMon = 0
   r_str_FecRec = ""
   r_dbl_IMPTOT = 0
   r_int_SinErr = 0
   r_int_ConErr = 0
   r_int_ErrorUpd = 0
   r_str_HorIni = Format(Time, "hhmmss")
   DoEvents
   
   r_int_NumFil = FreeFile
   Open p_NomFil For Input As r_int_NumFil
   
   'Leyendo Cabecera del Archivo
   Line Input #r_int_NumFil, r_str_Cadena
   
   If Left(r_str_Cadena, 2) = "01" Then
      r_str_NumCta = Mid(r_str_Cadena, 28, 18)
      r_str_FecRec = Mid(r_str_Cadena, 20, 8)
      
      Select Case Mid(r_str_Cadena, 17, 3)
         Case "USD": r_int_TipMon = 2
         Case "PEN": r_int_TipMon = 1
      End Select
   End If
   
   '*** inicializa log
   moddat_g_int_CntErr = 0
   g_str_Parame = "USP_CRE_PROPAGCAB ("
   g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 1 & ", "
   g_str_Parame = g_str_Parame & "'" & p_NomFil & "', "
   g_str_Parame = g_str_Parame & "'" & Dir(p_NomFil, vbArchive) & "', "
   g_str_Parame = g_str_Parame & r_str_FecRec & ", "
   g_str_Parame = g_str_Parame & CInt(r_str_CodBco) & ", "
   g_str_Parame = g_str_Parame & "'" & r_str_NumCta & "', "
   g_str_Parame = g_str_Parame & CInt(r_int_TipMon) & " , "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & 1 & " ) "
   
   Do While (moddat_g_int_CntErr = 0)
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            moddat_g_int_CntErr = 1
            Close #r_int_NumFil
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      Else
         moddat_g_int_CntErr = 1
      End If
   Loop
      
   g_rst_Princi.MoveFirst
   r_int_NumPro = g_rst_Princi!CORRELATIVO
   
   r_int_FlgErr = 0
   r_int_NFila = 0
   r_str_OpeAux = ""
   ReDim l_arr_GasCie(0)
   
   Do While Not EOF(r_int_NumFil)
      Line Input #r_int_NumFil, r_str_Cadena
      DoEvents
      
      r_int_CntErr = 0
      r_int_ErrorUpd = 0
      
      If Left(r_str_Cadena, 2) = "02" Then
         r_str_NomCli = Trim(Mid(r_str_Cadena, 3, 30))
         r_str_DocIde = Format(Trim(Mid(r_str_Cadena, 33, 13)), "#############")
         r_int_TipDoc = CInt(Mid(r_str_DocIde, 1, 1))
         r_str_NumDoc = Trim(Mid(r_str_DocIde, 2))
         r_str_NumOpe = Trim(Mid(r_str_Cadena, 46, 20))
         r_int_TipPag = CInt(Mid(r_str_Cadena, 66, 2))
         r_int_NumCuo = CInt(Mid(r_str_Cadena, 68, 3))
         r_str_FecPag = Mid(r_str_Cadena, 136, 8)
         r_str_NumMov = Mid(r_str_Cadena, 130, 6)
         r_dbl_ImpDep = CDbl(Mid(r_str_Cadena, 96, 13) & "." & Mid(r_str_Cadena, 109, 2))
         r_str_OfiPag = Mid(r_str_Cadena, 126, 4)
         
         Select Case Mid(r_str_Cadena, 144, 2)
            Case "01":  r_str_ForPag = "01 - EFECTIVO"
            Case "02":  r_str_ForPag = "02 - CARGO EN CUENTA"
            Case "03":  r_str_ForPag = "03 - CHEQUE BBVA O PROPIO BANCO"
            Case "04":  r_str_ForPag = "04 - CHEQUE OTRO BANCO"
            Case "05":  r_str_ForPag = "05 - CHEQUE REMESA"
            Case "06":  r_str_ForPag = "06 - MIXTO"
            Case "07":  r_str_ForPag = "07 - TARJETA DE CREDITO"
            Case "08":  r_str_ForPag = "08 - TELEPROCESO"
         End Select
         
         Select Case Mid(r_str_Cadena, 146, 2)
            Case "01":  r_str_CanPag = "01 - TERMINAL FINANCIERO / VENTANILLA"
            Case "02":  r_str_CanPag = "02 - CAJERO AUTOMATICO"
            Case "03":  r_str_CanPag = "03 - BANCA TELEFONICA"
            Case "04":  r_str_CanPag = "04 - BANCA POR INTERNET"
            Case "05":  r_str_CanPag = "05 - PAGOS EN LINEA"
         End Select
         
         Select Case r_int_TipPag
            Case 1   'GASTOS ADMINISTRATIVOS
               If Not ff_GasAdm(r_str_NumOpe, r_int_TipDoc, r_str_NumDoc, r_str_CodBco, r_str_FecPag, r_str_NumCta, r_str_NumMov, r_int_TipMon, r_dbl_ImpDep, r_str_FecRec, r_str_OfiPag, r_str_ForPag, r_str_CanPag, r_str_CadErr) Then
                  r_int_CntErr = r_int_CntErr + 1
                  r_int_ConErr = r_int_ConErr + 1
                  r_str_Situac = 1
               Else
                  r_int_SinErr = r_int_SinErr + 1
                  r_str_Situac = 2
               End If
               
            Case 2   'PAGO DE CUOTAS
               If Not (r_int_FlgErr = 1 And r_str_OpeAux = r_str_NumOpe) Then
                  If Not ff_CuoHip(r_str_NumOpe, r_int_TipDoc, r_str_NumDoc, r_str_CodBco, r_str_FecPag, r_str_NumCta, r_str_NumMov, r_int_TipMon, r_dbl_ImpDep, r_str_FecRec, r_str_OfiPag, r_str_ForPag, r_str_CanPag, r_int_NumCuo, r_str_CadErr) Then
                     'mensaje de cuotas ya pagadas no se consideran error
                     If r_str_CadErr = "La cuota ya ha sido pagada." Then
                        r_int_FlgErr = 0
                     Else
                        r_int_FlgErr = 1
                     End If
                     r_str_OpeAux = r_str_NumOpe
                     r_int_CntErr = r_int_CntErr + 1
                     r_int_ConErr = r_int_ConErr + 1
                     r_str_Situac = 1
                  Else
                     r_int_FlgErr = 0
                     r_str_OpeAux = ""
                     r_int_SinErr = r_int_SinErr + 1
                     r_str_Situac = 2
                  End If
               End If
               
            Case 5   'PAGO DE PLAN AHORRO
               If Not ff_PlaAho(r_str_NumOpe, r_int_TipDoc, r_str_NumDoc, r_str_CodBco, r_str_FecPag, r_str_NumCta, r_str_NumMov, r_int_TipMon, r_dbl_ImpDep, r_str_FecRec, r_str_OfiPag, r_str_ForPag, r_str_CanPag, r_int_NumCuo, r_str_CadErr) Then
                  r_int_CntErr = r_int_CntErr + 1
                  r_int_ConErr = r_int_ConErr + 1
                  r_str_Situac = 1
               Else
                  r_int_SinErr = r_int_SinErr + 1
                  r_str_Situac = 2
               End If
               
         End Select
         
         r_int_NFila = r_int_NFila + 1
         
         '*** actualiza log
         g_str_Parame = "USP_CRE_PROPAGDET ("
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & r_int_NumPro & ", "
         g_str_Parame = g_str_Parame & r_int_NFila & ", "
         g_str_Parame = g_str_Parame & r_str_Situac & " , "
         g_str_Parame = g_str_Parame & "'" & r_str_NumOpe & "' , "
         g_str_Parame = g_str_Parame & "'" & r_str_NumDoc & "' , "
         g_str_Parame = g_str_Parame & "'" & r_str_NomCli & "', "
         g_str_Parame = g_str_Parame & r_int_TipPag & ", "
         g_str_Parame = g_str_Parame & r_str_FecPag & ", "
         g_str_Parame = g_str_Parame & r_dbl_ImpDep & ", "
         g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
         g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
         
         Do While (r_int_ErrorUpd = 0)
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  r_int_ErrorUpd = 1
               Else
                  r_int_ErrorUpd = 0
               End If
            Else
               r_int_ErrorUpd = 1
            End If
         Loop
         
      End If
   Loop
   
   Close #r_int_NumFil
   
   r_str_HorFin = Format(Time, "hhmmss")
   
   If Left(r_str_Cadena, 2) = "03" Then
      r_int_NumReg = CInt(Mid(r_str_Cadena, 5, 7))
      r_dbl_IMPTOT = CDbl(Mid(r_str_Cadena, 12, 15)) / 100
      
      '*** finaliza log
      g_str_Parame = "USP_CRE_PROPAGCAB ("
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & r_int_NumPro & ", "
      g_str_Parame = g_str_Parame & 2 & ", "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & r_str_FecRec & ", "
      g_str_Parame = g_str_Parame & CInt(r_str_CodBco) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_NumCta & "', "
      g_str_Parame = g_str_Parame & CInt(r_int_TipMon) & " , "
      g_str_Parame = g_str_Parame & r_int_NumReg & ", "
      g_str_Parame = g_str_Parame & r_dbl_IMPTOT & ", "
      g_str_Parame = g_str_Parame & r_int_NumFil & ", "
      g_str_Parame = g_str_Parame & r_int_ConErr & ", "
      g_str_Parame = g_str_Parame & r_int_SinErr & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & 0 & " ) "
      
      Do While (r_int_ErrorUpd = 0)
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               r_int_ErrorUpd = 1
            Else
               r_int_ErrorUpd = 0
            End If
         Else
            r_int_ErrorUpd = 1
         End If
      Loop
   
   End If
   
   Call fs_Correo_GasCie(cmb_CodBan.Text, moddat_gf_Consulta_ParDes("204", CStr(r_int_TipMon)))
   
   'Generando Reporte
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.SelectionFormula = "{CRE_PROPAGCAB.PAGCAB_FECPRO} = " & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & " AND {CRE_PROPAGCAB.PAGCAB_NUMPRO} = " & r_int_NumPro & " "
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_COBBCO_01.RPT"
   crp_Imprim.Action = 1
End Sub

Private Function ff_GasAdm(ByVal p_NumSol As String, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_CodBco As String, ByVal p_FecPag As String, ByVal p_CtaBan As String, ByVal p_NumCom As String, ByVal p_TipMon As Integer, ByVal p_TotDep As Double, ByVal p_FecRec As String, ByVal p_OfiPag As String, ByVal p_ForPag As String, ByVal p_CanPag As String, ByRef p_CadErr As String) As Integer
Dim r_rst_Princi     As ADODB.Recordset
Dim r_dbl_TotGas     As Double
Dim r_int_CodGas     As Integer
Dim r_dbl_ImpGas     As Double
Dim r_dbl_ImpITF     As Double
Dim r_dbl_TotPag     As Double
Dim r_int_MonOpe     As Integer
Dim r_str_Operac     As String
Dim r_lng_NumMov     As Long
Dim r_str_CodPrd     As String
Dim r_str_ConHip     As String
Dim r_str_EjeSeg     As String
Dim r_int_TipDoc     As Integer
Dim r_str_NumDoc     As String
Dim r_str_NomCli     As String
Dim r_int_CodIns     As Integer
   
   ff_GasAdm = False
   l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
   p_CadErr = ""
   
   'Valida Solicitud
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = '" & p_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      p_CadErr = "Error al acceder a la Base de Datos para validar la solicitud"
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      p_CadErr = "No se encontró la solicitud en la Base de Datos"
      Exit Function
   End If

   r_str_CodPrd = r_rst_Princi!SOLMAE_CODPRD
   r_int_MonOpe = r_rst_Princi!SOLMAE_TIPMON
   r_int_TipDoc = r_rst_Princi!SOLMAE_TITTDO
   r_str_NumDoc = Trim(r_rst_Princi!SOLMAE_TITNDO & "")
   r_str_ConHip = Trim(r_rst_Princi!SOLMAE_CONHIP & "")
   r_str_EjeSeg = Trim(r_rst_Princi!SOLMAE_EJESEG & "")
   r_str_NomCli = moddat_gf_Buscar_NomCli(CStr(r_int_TipDoc), Trim(r_str_NumDoc))
   r_int_CodIns = r_rst_Princi!SOLMAE_CODINS
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   If p_TipMon <> r_int_MonOpe Then
      p_CadErr = "La moneda de pago no coincide con la moneda de solicitud"
      Exit Function
   End If
   
   'Validando Total de Gastos de Cierre
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_GASADM "
   g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & p_NumSol & "' "
   g_str_Parame = g_str_Parame & "   AND GASADM_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      p_CadErr = "Error al acceder a la Base de Datos para validar el gasto de cierre (1)"
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      p_CadErr = "El gasto de cierre ya fue pagado."
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Exit Function
   End If
   
   r_dbl_TotGas = 0
   r_rst_Princi.MoveFirst
   Do While Not r_rst_Princi.EOF
      r_dbl_ImpGas = r_rst_Princi!GASADM_IMPORT
      r_dbl_TotGas = r_dbl_TotGas + r_dbl_ImpGas
      r_rst_Princi.MoveNext
   Loop
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   r_dbl_ImpITF = gf_NueImp_Numero(gf_Truncar_Numero(r_dbl_TotGas * (l_dbl_PorITF / 100), 2))
   r_dbl_TotPag = CDbl(Format(r_dbl_TotGas + r_dbl_ImpITF, "###0.00"))
   
   If CDbl(r_dbl_TotPag) <> CDbl(p_TotDep) Then
      p_CadErr = "El importe depositado no coincide con el importe adeudado."
      Exit Function
   End If
   
   'Obteniendo Código de Operación Contable
   r_str_Operac = moddat_gf_Consulta_Operac(r_str_CodPrd, "211")
   r_str_Operac = CStr(p_TipMon) & Right(r_str_Operac, 5)
   
   'Obteniendo Número de Movimiento
   r_lng_NumMov = opecaj_gf_Genera_NumMov()
   
   'Registrando Movimiento
   If Not opecaj_gf_Inserta_CajMov(modgen_g_str_CodUsu, "1101", p_NumSol, "", p_TipDoc, p_NumDoc, p_CodBco, p_FecPag, p_CtaBan, p_NumCom, p_TipMon, r_dbl_TotGas, 0, modgen_g_str_CodSuc, 0, 0, 0, l_dbl_PorITF, r_dbl_ImpITF, p_TotDep, 0, "0", r_str_Operac, r_lng_NumMov, 2, p_FecRec, p_OfiPag, p_ForPag, p_CanPag, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0) Then
      p_CadErr = "Error al insertar movimiento para la contabilidad - gastos de cierre"
      Exit Function
   End If
   
   'Buscar Gastos de Cierre
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_GASADM "
   g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & p_NumSol & "' "
   g_str_Parame = g_str_Parame & "   AND GASADM_SITUAC = 2 "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      p_CadErr = "Error al acceder a la Base de Datos para validar el gasto de cierre (2)"
      Exit Function
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      Do While Not r_rst_Princi.EOF
         r_int_CodGas = r_rst_Princi!GASADM_CODGAS
         r_dbl_ImpGas = r_rst_Princi!GASADM_IMPORT
      
         'Actualizando Pago en Tabla de Gasto Administrativo
         If Not opecaj_gf_Pago_GasAdm(p_NumSol, r_int_CodGas, p_TipMon, r_dbl_ImpGas, l_dbl_PorITF, p_FecPag, r_str_Operac) Then
            p_CadErr = "Error al registrar el pago del gastos de cierre"
            Exit Function
         End If
         
         r_rst_Princi.MoveNext
      Loop
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'Actualizando en Seguimiento de Tasacion Pago de Gastos Administrativos
   If Not moddat_gf_Inserta_SegDet(p_NumSol, r_int_CodIns, 25, 0, "", 0, 0) Then
      p_CadErr = "Error al actualizar seguimientos"
      Exit Function
   End If
   
   ReDim Preserve l_arr_GasCie(UBound(l_arr_GasCie) + 1)
   l_arr_GasCie(UBound(l_arr_GasCie)).Genera_NumSol = gf_Formato_NumSol(p_NumSol)
   l_arr_GasCie(UBound(l_arr_GasCie)).Genera_TipDoc = r_int_TipDoc
   l_arr_GasCie(UBound(l_arr_GasCie)).Genera_NumDoc = r_str_NumDoc
   l_arr_GasCie(UBound(l_arr_GasCie)).Genera_NomCli = r_str_NomCli
   l_arr_GasCie(UBound(l_arr_GasCie)).Genera_ConHip = r_str_ConHip
   l_arr_GasCie(UBound(l_arr_GasCie)).Genera_EjeSeg = r_str_EjeSeg
   l_arr_GasCie(UBound(l_arr_GasCie)).Genera_CodIns = r_int_CodIns
   
   p_CadErr = "Gasto de cierre pagado satisfactoriamente"
   ff_GasAdm = True
End Function

Private Function ff_CuoHip(ByVal p_NumOpe As String, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_CodBco As String, ByVal p_FecPag As String, ByVal p_CtaBan As String, ByVal p_NumCom As String, ByVal p_TipMon As Integer, ByVal p_TotDep As Double, ByVal p_FecRec As String, ByVal p_OfiPag As String, ByVal p_ForPag As String, ByVal p_CanPag As String, ByVal p_NumCuo As Integer, ByRef p_CadErr As String) As Integer
Dim r_rst_Princi     As ADODB.Recordset
Dim r_dbl_TotCuo     As Double
Dim r_dbl_ImpITF     As Double
Dim r_dbl_TotPag     As Double
Dim r_int_MonOpe     As Integer
Dim r_str_Operac     As String
Dim r_lng_NumMov     As Long
Dim r_str_CodPrd     As String
Dim r_int_SitCre     As Integer
Dim r_int_SitAnt     As Integer
Dim r_int_Situac     As Integer
Dim r_int_CuoPen     As Integer
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_SegDes     As Double
Dim r_dbl_SegInm     As Double
Dim r_dbl_Portes     As Double
Dim r_dbl_CapBBP     As Double
Dim r_dbl_IntBBP     As Double
Dim r_dbl_IntCom     As Double
Dim r_dbl_IntMor     As Double
Dim r_dbl_GasCob     As Double
Dim r_dbl_OtrGas     As Double
Dim r_str_PrxVct     As String
Dim r_str_CodSub     As String
   
   ff_CuoHip = False
   p_CadErr = ""
   
   'Valida Operacion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & p_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      p_CadErr = "Error al acceder a la Base de Datos para validar la Operación"
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      p_CadErr = "No se encontró la operación en la Base de Datos"
      Exit Function
   End If
   
'   If r_rst_Princi!HIPMAE_SITUAC = 6 Then
'      r_rst_Princi.Close
'      Set r_rst_Princi = Nothing
'      p_CadErr = "El pago corresponde a un cliente transferido"
'      Exit Function
'   End If
   
   r_str_CodPrd = r_rst_Princi!HIPMAE_CODPRD       'Codigo producto
   r_int_MonOpe = r_rst_Princi!HIPMAE_MONEDA       'Codigo moneda
   r_int_SitCre = r_rst_Princi!HIPMAE_SITCRE       'Situación de Crédito SBS
   r_int_SitAnt = r_rst_Princi!HIPMAE_SITANT       'Situación Anterior
   r_int_Situac = r_rst_Princi!HIPMAE_SITUAC       'Situación de Crédito miCasita
   r_int_CuoPen = r_rst_Princi!HIPMAE_CUOPEN       'Cuotas Pendientes
   r_str_CodSub = r_rst_Princi!HIPMAE_CODSUB       'Codigo subproducto
   
   'Obteniendo ITF
   If r_rst_Princi!HIPMAE_INDITF = 2 Then
      l_dbl_PorITF = 0
   Else
      l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   If p_TipMon <> r_int_MonOpe Then
      p_CadErr = "La moneda de pago no coincide con la moneda de la operación"
      Exit Function
   End If
   
   'Validando Cuota / Calculando Total Cuota e ITF
   r_dbl_TotCuo = 0
   r_dbl_ImpITF = 0
   r_dbl_TotPag = 0
      
   'Valida gastos de cobranza, actualiza el campo HIPCUO_GASCOB si fuera el caso
   Call ValidaGastosCobranza(p_NumOpe, p_NumCuo, p_FecPag, p_TotDep, r_str_CodPrd, r_str_CodSub, l_dbl_PorITF)
   '******************************************************************************************************************
   
   'Valida interes moratorio, actualiza el campo HIPCUO_INTMOR si fuera el caso
   Call ValidaInteresMoratio(p_NumOpe, p_NumCuo, p_TotDep, l_dbl_PorITF)
   '******************************************************************************************************************
   
   'Valida si quedan cuotas anteriores pendientes de pago
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS NUMCUO_PEND "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO < " & CStr(p_NumCuo)
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      p_CadErr = "Error al acceder a la Base de Datos para validar cuotas anteriores pendientes de pago"
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      p_CadErr = "La cuota ya ha sido pagada."
      Exit Function
   End If
   
   r_rst_Princi.MoveFirst
   If r_rst_Princi!NUMCUO_PEND > 0 Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      p_CadErr = "El cliente tiene cuotas anteriores pendientes de pago."
      Exit Function
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'Valida si la cuota ya fue pagada
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO = " & CStr(p_NumCuo)
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      p_CadErr = "Error al acceder a la Base de Datos para validar la cuota"
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      p_CadErr = "La cuota ya ha sido pagada."
      Exit Function
   End If
   
   r_dbl_Capita = r_rst_Princi!HIPCUO_CAPITA - r_rst_Princi!HIPCUO_CAPPAG
   r_dbl_Intere = r_rst_Princi!HIPCUO_INTERE - r_rst_Princi!HIPCUO_INTPAG
   r_dbl_SegDes = r_rst_Princi!HIPCUO_DESORG - r_rst_Princi!HIPCUO_DESPAG
   r_dbl_SegInm = r_rst_Princi!HIPCUO_VIVORG - r_rst_Princi!HIPCUO_VIVPAG
   r_dbl_Portes = r_rst_Princi!HIPCUO_OTRORG - r_rst_Princi!HIPCUO_OTRPAG
   r_dbl_CapBBP = r_rst_Princi!HIPCUO_CAPBBP - r_rst_Princi!HIPCUO_CBPPAG
   r_dbl_IntBBP = r_rst_Princi!HIPCUO_INTBBP - r_rst_Princi!HIPCUO_IBPPAG
   r_dbl_IntCom = r_rst_Princi!HIPCUO_INTCOM - r_rst_Princi!HIPCUO_ICOPAG
   r_dbl_IntMor = r_rst_Princi!HIPCUO_INTMOR - r_rst_Princi!HIPCUO_IMOPAG
   r_dbl_GasCob = r_rst_Princi!HIPCUO_GASCOB - r_rst_Princi!HIPCUO_GCOPAG
   r_dbl_OtrGas = r_rst_Princi!HIPCUO_OTRGAS - r_rst_Princi!HIPCUO_OTGPAG
   r_dbl_TotCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_SegDes + r_dbl_SegInm + r_dbl_Portes + r_dbl_CapBBP + r_dbl_IntBBP + r_dbl_IntCom + r_dbl_IntMor + r_dbl_GasCob + r_dbl_OtrGas
   r_dbl_ImpITF = gf_NueImp_Numero(gf_Truncar_Numero(r_dbl_TotCuo * (l_dbl_PorITF / 100), 2))
   r_dbl_TotPag = CDbl(Format(r_dbl_TotCuo + r_dbl_ImpITF, "###0.00"))
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   If CDbl(r_dbl_TotPag) <> CDbl(p_TotDep) Then
      p_CadErr = "El Importe depositado no coincide con el Importe adeudado."
      Exit Function
   End If
   
   'Para obtener Próximo Vencimiento
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      p_CadErr = "Error al acceder a la Base de Datos para validar el proximo vencimiento"
      Exit Function
   End If
   
   If r_rst_Princi.BOF Or r_rst_Princi.EOF Then
      r_int_CuoPen = 0
      r_int_Situac = 9
      r_str_PrxVct = "0"
   Else
      r_rst_Princi.MoveFirst
      DoEvents
      r_rst_Princi.MoveNext
      
      If r_rst_Princi.BOF Or r_rst_Princi.EOF Then
         r_int_CuoPen = 0
         r_int_Situac = 9
         r_str_PrxVct = "0"
      Else
         r_int_CuoPen = r_int_CuoPen - 1
         r_str_PrxVct = CStr(r_rst_Princi!HIPCUO_FECVCT)
      End If
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'Obteniendo Código de Operación Contable
   r_str_Operac = moddat_gf_Consulta_Operac(r_str_CodPrd, "033")
   r_str_Operac = CStr(p_TipMon) & Right(r_str_Operac, 5)
   
   'Obteniendo Número de Movimiento
   r_lng_NumMov = opecaj_gf_Genera_NumMov()
   
   'Registrando Movimiento
   If Not opecaj_gf_Inserta_CajMov(modgen_g_str_CodUsu, "1102", p_NumOpe, "", p_TipDoc, p_NumDoc, p_CodBco, p_FecPag, p_CtaBan, p_NumCom, p_TipMon, r_dbl_TotCuo, 0, modgen_g_str_CodSuc, 0, 0, 0, l_dbl_PorITF, r_dbl_ImpITF, p_TotDep, 0, "0", r_str_Operac, r_lng_NumMov, 2, p_FecRec, p_OfiPag, p_ForPag, p_CanPag, r_dbl_Capita, r_dbl_Intere, r_dbl_SegDes, r_dbl_SegInm, r_dbl_Portes, r_dbl_CapBBP, r_dbl_IntBBP, r_dbl_IntMor, r_dbl_IntCom, r_dbl_GasCob, r_dbl_OtrGas, "", 0) Then
      p_CadErr = "Error al insertar movimiento para la contabilidad - cuotas"
      Exit Function
   End If
   
   'Pagando Cuota
   If Not opecaj_gf_Pago_Cuotas(p_NumOpe, p_NumCuo, p_FecPag, r_dbl_TotCuo, r_dbl_Capita, r_dbl_Intere, r_dbl_SegDes, r_dbl_SegInm, r_dbl_Portes, r_dbl_IntCom, r_dbl_IntMor, r_dbl_GasCob, r_dbl_OtrGas, 0, 0, r_int_SitCre, r_str_Operac, r_lng_NumMov, 1, r_str_PrxVct, r_int_CuoPen, r_int_Situac, r_int_SitAnt, 2, r_dbl_CapBBP, r_dbl_IntBBP) Then
      p_CadErr = "Error al registrar el pago de cuota"
      Exit Function
   End If
   
   p_CadErr = "Cuota pagada satisfactoriamente"
   ff_CuoHip = True
End Function

Private Function ff_PlaAho(ByVal p_NumOpe As String, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_CodBco As String, ByVal p_FecPag As String, ByVal p_CtaBan As String, ByVal p_NumCom As String, ByVal p_TipMon As Integer, ByVal p_TotDep As Double, ByVal p_FecRec As String, ByVal p_OfiPag As String, ByVal p_ForPag As String, ByVal p_CanPag As String, ByVal p_NumCuo As Integer, ByRef p_CadErr As String) As Integer
Dim r_rst_Princi     As ADODB.Recordset
Dim r_dbl_TotCuo     As Double
Dim r_dbl_ImpITF     As Double
Dim r_dbl_TotPag     As Double
Dim r_int_MonOpe     As Integer
Dim r_str_Operac     As String
Dim r_lng_NumMov     As Long
Dim r_str_CodPrd     As String
Dim r_int_SitCre     As Integer
Dim r_int_SitAnt     As Integer
Dim r_str_Situac     As String
Dim r_int_CuoPen     As Integer
Dim r_str_PrxVct     As String
Dim r_str_CodSub     As String
Dim r_dbl_Capita     As Double
Dim r_dbl_SalCap     As Double
Dim r_int_NumCuo     As Integer
Dim r_str_FecVct     As String
   
   ff_PlaAho = False
   p_CadErr = ""
   
   'Validar Plan de Ahorro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_AHOMAE "
   g_str_Parame = g_str_Parame & " WHERE AHOMAE_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "   AND AHOMAE_NUMDOC = '" & Trim(p_NumDoc) & "' "
   g_str_Parame = g_str_Parame & "   AND AHOMAE_NUMERO = '" & Trim(p_NumOpe) & "' "
   g_str_Parame = g_str_Parame & "   AND AHOMAE_SITUAC = '2' "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      p_CadErr = "Error al acceder a la Base de Datos para validar el plan de ahorro"
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      p_CadErr = "No se encontró el plan de ahorro en la Base de Datos"
      Exit Function
   End If
   
   r_int_MonOpe = r_rst_Princi!AHOMAE_MONAHO
   r_int_NumCuo = r_rst_Princi!AHOMAE_NUMMES       'Nro. Cuota
   r_str_CodPrd = r_rst_Princi!AHOMAE_CODPRD       'Codigo producto
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   If p_TipMon <> r_int_MonOpe Then
      p_CadErr = "La moneda de pago no coincide con la moneda del plan de ahorro"
      Exit Function
   End If
   
   'Para obtener valores de la cuota
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_AHOCUO "
   g_str_Parame = g_str_Parame & " WHERE AHOCUO_NUMERO = '" & Trim(p_NumOpe) & "' "
   g_str_Parame = g_str_Parame & "   AND AHOCUO_NUMCUO = " & p_NumCuo & "  "
   g_str_Parame = g_str_Parame & "ORDER BY AHOCUO_NUMCUO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      p_CadErr = "Error al acceder a la Base de Datos para buscar las cuotas del plan de ahorros"
      Exit Function
   End If
      
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      p_CadErr = "No se encontró las cuotas del plan de ahorro en la base de datos"
      Exit Function
   End If
      
   r_str_Situac = r_rst_Princi!AHOCUO_SITUAC       'Estado
   r_str_FecVct = Trim(r_rst_Princi!AHOCUO_FECVCT) 'Vencimiento
   r_dbl_TotCuo = r_rst_Princi!AHOCUO_CAPITA       'Total cuota
   r_dbl_Capita = r_rst_Princi!AHOCUO_CAPITA       'Total cuota
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
     
   If CInt(r_str_Situac) = 9 Then
      p_CadErr = "La cuota del plan de ahorros ya ha sido pagada."
      Exit Function
   End If
   
   'Obteniendo Código de Operación Contable
   r_str_Operac = moddat_gf_Consulta_Operac(r_str_CodPrd, "033")
   r_str_Operac = CStr(p_TipMon) & Right(r_str_Operac, 5)
   
   'Obteniendo Número de Movimiento
   r_lng_NumMov = opecaj_gf_Genera_NumMov()
   
   'Registrando Movimiento
   If Not opecaj_gf_Inserta_CajMov(modgen_g_str_CodUsu, "1105", p_NumOpe, "", p_TipDoc, p_NumDoc, p_CodBco, p_FecPag, p_CtaBan, p_NumCom, p_TipMon, r_dbl_TotCuo, 0, modgen_g_str_CodSuc, 0, 0, 0, 0, 0, p_TotDep, 0, "0", r_str_Operac, r_lng_NumMov, 2, p_FecRec, p_OfiPag, p_ForPag, p_CanPag, r_dbl_Capita, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0) Then
      p_CadErr = "Error al insertar movimiento para la contabilidad - plan de ahorros"
      Exit Function
   End If
               
   'Grabando Información del pago de cuota, cancelacion de la cuota y actualizacion del saldo pendiente y pagado en el maestro.
   g_str_Parame = "USP_CAN_PLAAHO_PAGO ("
   g_str_Parame = g_str_Parame & "'" & Trim(p_NumOpe) & "', "
   g_str_Parame = g_str_Parame & "" & p_NumCuo & ", "
   g_str_Parame = g_str_Parame & "" & r_str_FecVct & ","
   g_str_Parame = g_str_Parame & "" & p_FecPag & ", "
   g_str_Parame = g_str_Parame & "" & CDbl(p_TotDep) & ", "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   
   g_str_Parame = g_str_Parame & "" & CInt(r_int_NumCuo) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumCom) & ", "
   g_str_Parame = g_str_Parame & "'" & r_str_Operac & "') "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       p_CadErr = "Error al ejecutar el Procedimiento USP_CAN_PLAAHO_PAGO." & " - " & modgen_g_str_NomPlt
       Exit Function
   End If
   
   p_CadErr = "Cuota del plan de ahorro pagada satisfactoriamente"
   ff_PlaAho = True
End Function

Private Sub fs_Graba_Report(ByVal p_NumOpe As String, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_FecPag As String, ByVal p_Import As Double, ByVal p_TipPag As Integer, ByVal p_NumCuo As Integer, ByVal p_DesErr As String)
   'Grabando en DAO
   moddat_g_str_CadDAO = "SELECT * FROM RPT_COBBCO WHERE COBBCO_NUMOPE = ' '"
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
   moddat_g_rst_RecDAO("COBBCO_NUMOPE") = p_NumOpe
   moddat_g_rst_RecDAO("COBBCO_DOCIDE") = CStr(p_TipDoc) & "-" & Trim(p_NumDoc & "")
   moddat_g_rst_RecDAO("COBBCO_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(p_TipDoc), Trim(p_NumDoc))
   moddat_g_rst_RecDAO("COBBCO_FECPAG") = gf_FormatoFecha(p_FecPag)
   moddat_g_rst_RecDAO("COBBCO_IMPORT") = p_Import
   moddat_g_rst_RecDAO("COBBCO_TIPPAG") = moddat_gf_Consulta_ParDes("240", CStr(p_TipPag))
   moddat_g_rst_RecDAO("COBBCO_NUMCUO") = p_NumCuo
   moddat_g_rst_RecDAO("COBBCO_DESERR") = p_DesErr
   moddat_g_rst_RecDAO.Update
   DoEvents
   
   moddat_g_rst_RecDAO.Close
   DoEvents
End Sub

Private Sub fs_Correo_GasCie(ByVal p_CodBan As String, ByVal p_Moneda As String)
Dim r_int_Contad     As Integer
Dim r_str_Cadena     As String
   
   If UBound(l_arr_GasCie) = 0 Then
      Exit Sub
   End If
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Enviando Correo Electrónico
   modgen_g_str_Mail_Asunto = "PAGO DE GASTOS DE CIERRE - " & Trim(p_CodBan) & " - " & Trim(p_Moneda) & " (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   modgen_g_str_Mail_Mensaj = ""
   
   For r_int_Contad = 1 To UBound(l_arr_GasCie)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD   : " & l_arr_GasCie(r_int_Contad).Genera_NumSol & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE            : " & CStr(l_arr_GasCie(r_int_Contad).Genera_TipDoc) & "-" & l_arr_GasCie(r_int_Contad).Genera_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE        : " & l_arr_GasCie(r_int_Contad).Genera_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "CONSEJERO HIPOTECARIO : " & l_arr_GasCie(r_int_Contad).Genera_ConHip & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "EJECUTIVO SEGUIMIENTO : " & l_arr_GasCie(r_int_Contad).Genera_EjeSeg & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   
      'Usuario de Seguimiento
      r_str_Cadena = moddat_gf_Buscar_DirEle_Codigo(l_arr_GasCie(r_int_Contad).Genera_EjeSeg)
   
      If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
         ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
         moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
      End If
   
      'Consejero Hipotecario
      r_str_Cadena = moddat_gf_Buscar_DirEle_Codigo(l_arr_GasCie(r_int_Contad).Genera_ConHip)
   
      If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
         ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
         moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
      End If
   Next r_int_Contad
   
   'Jefe de Seguimiento
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(130)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Ventas
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(120)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director Comercial
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(100)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Operaciones
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(220)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Evaluador de Operaciones
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(221)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
End Sub

Private Sub ValidaInteresMoratio(ByVal p_Operac As String, ByVal p_Cuota As Integer, ByVal p_Depositado As Double, ByVal p_ITF As Double)
Dim r_rst_IntMor     As ADODB.Recordset
Dim r_str_CadSql     As String
Dim r_dbl_MtoCap     As Double
Dim r_dbl_MtoInt     As Double
Dim r_dbl_MtoSDg     As Double
Dim r_dbl_MtoSIn     As Double
Dim r_dbl_MtoPor     As Double
Dim r_dbl_MtpCBP     As Double
Dim r_dbl_MtoIBP     As Double
Dim r_dbl_MtoICo     As Double
Dim r_dbl_MtoIMo     As Double
Dim r_dbl_MtoGCo     As Double
Dim r_dbl_MtoOGa     As Double
Dim r_dbl_MtoTot     As Double
Dim r_dbl_MtoITF     As Double
Dim r_dbl_TotCuo     As Double
Dim r_dbl_MtoDif     As Double
   
   r_str_CadSql = ""
   r_str_CadSql = r_str_CadSql & "SELECT * "
   r_str_CadSql = r_str_CadSql & "  FROM CRE_HIPCUO "
   r_str_CadSql = r_str_CadSql & " WHERE HIPCUO_NUMOPE = '" & p_Operac & "' "
   r_str_CadSql = r_str_CadSql & "   AND HIPCUO_NUMCUO = " & CStr(p_Cuota)
   r_str_CadSql = r_str_CadSql & "   AND HIPCUO_TIPCRO = 1 "
   r_str_CadSql = r_str_CadSql & "   AND HIPCUO_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(r_str_CadSql, r_rst_IntMor, 3) Then
      Exit Sub
   End If
   
   If r_rst_IntMor.BOF And r_rst_IntMor.EOF Then
      r_rst_IntMor.Close
      Set r_rst_IntMor = Nothing
      Exit Sub
   End If
   
   r_dbl_MtoCap = r_rst_IntMor!HIPCUO_CAPITA - r_rst_IntMor!HIPCUO_CAPPAG
   r_dbl_MtoInt = r_rst_IntMor!HIPCUO_INTERE - r_rst_IntMor!HIPCUO_INTPAG
   r_dbl_MtoSDg = r_rst_IntMor!HIPCUO_DESORG - r_rst_IntMor!HIPCUO_DESPAG
   r_dbl_MtoSIn = r_rst_IntMor!HIPCUO_VIVORG - r_rst_IntMor!HIPCUO_VIVPAG
   r_dbl_MtoPor = r_rst_IntMor!HIPCUO_OTRORG - r_rst_IntMor!HIPCUO_OTRPAG
   r_dbl_MtpCBP = r_rst_IntMor!HIPCUO_CAPBBP - r_rst_IntMor!HIPCUO_CBPPAG
   r_dbl_MtoIBP = r_rst_IntMor!HIPCUO_INTBBP - r_rst_IntMor!HIPCUO_IBPPAG
   r_dbl_MtoICo = r_rst_IntMor!HIPCUO_INTCOM - r_rst_IntMor!HIPCUO_ICOPAG
   r_dbl_MtoIMo = r_rst_IntMor!HIPCUO_INTMOR - r_rst_IntMor!HIPCUO_IMOPAG
   r_dbl_MtoGCo = r_rst_IntMor!HIPCUO_GASCOB - r_rst_IntMor!HIPCUO_GCOPAG
   r_dbl_MtoOGa = r_rst_IntMor!HIPCUO_OTRGAS - r_rst_IntMor!HIPCUO_OTGPAG
   
   r_dbl_MtoTot = r_dbl_MtoCap + r_dbl_MtoInt + r_dbl_MtoSDg + r_dbl_MtoSIn + r_dbl_MtoPor + r_dbl_MtpCBP + r_dbl_MtoIBP + r_dbl_MtoICo + r_dbl_MtoIMo + r_dbl_MtoGCo + r_dbl_MtoOGa
   r_dbl_MtoITF = gf_NueImp_Numero(gf_Truncar_Numero(r_dbl_MtoTot * (p_ITF / 100), 2))
   r_dbl_TotCuo = CDbl(Format(r_dbl_MtoTot + r_dbl_MtoITF, "###0.00"))
   
   r_rst_IntMor.Close
   Set r_rst_IntMor = Nothing

   If CDbl(r_dbl_TotCuo) <> CDbl(p_Depositado) Then
      r_dbl_MtoDif = CDbl(Format(r_dbl_TotCuo - p_Depositado, "####0.00"))
      If (r_dbl_MtoDif > 0) And (r_dbl_MtoDif <= r_dbl_MtoIMo) Then
      
         'Actualizando en CRE_HIPCUO
         r_str_CadSql = ""
         r_str_CadSql = r_str_CadSql & "UPDATE CRE_HIPCUO "
         r_str_CadSql = r_str_CadSql & "   SET HIPCUO_INTMOR = HIPCUO_INTMOR - " & Format(r_dbl_MtoDif, "#######0.00") & " "
         r_str_CadSql = r_str_CadSql & " WHERE HIPCUO_NUMOPE = '" & p_Operac & "' "
         r_str_CadSql = r_str_CadSql & "   AND HIPCUO_NUMCUO = " & CStr(p_Cuota)
         r_str_CadSql = r_str_CadSql & "   AND HIPCUO_TIPCRO = 1 "
         r_str_CadSql = r_str_CadSql & "   AND HIPCUO_SITUAC = 2 "
         
         If Not gf_EjecutaSQL(r_str_CadSql, r_rst_IntMor, 2) Then
            MsgBox "Error al actualzar en CRE_HIPCUO (Operación) : " & p_Operac, vbInformation, modgen_g_str_NomPlt
         End If
      End If
   End If

End Sub

Private Sub ValidaGastosCobranza(ByVal p_NumOpe As String, ByVal p_NumCuo As Integer, ByVal p_FecPag As String, ByVal p_TotDep As Double, ByVal p_CodPrd As String, ByVal p_SubPrd As String, ByVal p_PorITF As Double)
Dim r_rst_Princi     As ADODB.Recordset
Dim r_rst_GasCob     As ADODB.Recordset
Dim r_dbl_TotCuo     As Double
Dim r_dbl_ImpITF     As Double
Dim r_dbl_TotPag     As Double
Dim r_dbl_TotPro     As Double
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_SegDes     As Double
Dim r_dbl_SegInm     As Double
Dim r_dbl_Portes     As Double
Dim r_dbl_CapBBP     As Double
Dim r_dbl_IntBBP     As Double
Dim r_dbl_IntCom     As Double
Dim r_dbl_IntMor     As Double
Dim r_dbl_GasCob     As Double
Dim r_dbl_OtrGas     As Double
Dim r_int_DiaAtrPag  As Integer
Dim r_int_DiaAtrPro  As Integer
Dim r_dbl_CobrzaPag  As Double
Dim r_dbl_CobrzaPro  As Double
Dim r_dbl_Cobrza     As Double
   
    r_dbl_Cobrza = 0
    r_dbl_CobrzaPro = 0
    r_dbl_CobrzaPag = 0
   
    ' Obtiene datos de la cuota
    g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
    g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
    g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
    g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
    g_str_Parame = g_str_Parame & "HIPCUO_NUMCUO = " & CStr(p_NumCuo)
    
    If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
    End If
    
    If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Exit Sub
    End If
    
    ' Comparacion con fecha de pago del archivo
    r_int_DiaAtrPag = CInt(CDate(gf_FormatoFecha(p_FecPag)) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!HIPCUO_FECVCT))))
    
    ' Leyendo Gastos de Cobranzas según días de atraso -- Fecha de pago
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " SELECT GASCOB_IMPORT "
    g_str_Parame = g_str_Parame & "   FROM OPE_GASCOB_CAB "
    g_str_Parame = g_str_Parame & "        INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = '" & p_NumOpe & "'"
    g_str_Parame = g_str_Parame & "                             AND GASCOBCAB_FECINI <= HIPMAE_FECDES "
    g_str_Parame = g_str_Parame & "                             AND GASCOBCAB_FECFIN >= HIPMAE_FECDES "
    g_str_Parame = g_str_Parame & "        INNER JOIN OPE_GASCOB ON GASCOBCAB_CODRAN = GASCOB_CODRAN  "
    g_str_Parame = g_str_Parame & "                             AND GASCOB_DIAINI <= " & CStr(r_int_DiaAtrPag) & " "
    g_str_Parame = g_str_Parame & "                             AND GASCOB_DIAFIN >= " & CStr(r_int_DiaAtrPag) & " "
    g_str_Parame = g_str_Parame & "  WHERE GASCOBCAB_CODPRD = '" & p_CodPrd & "' "
    g_str_Parame = g_str_Parame & "    AND GASCOBCAB_CODSUB = '" & p_SubPrd & "' "
    g_str_Parame = g_str_Parame & "    AND GASCOBCAB_CODRAN <> 0 "
    
    If Not gf_EjecutaSQL(g_str_Parame, r_rst_GasCob, 3) Then
       MsgBox "Error al leer OPE_GASCOB (Operación) : " & p_NumOpe, vbInformation, modgen_g_str_NomPlt
    End If
    
    If Not (r_rst_GasCob.BOF And r_rst_GasCob.EOF) Then
       r_dbl_CobrzaPag = r_rst_GasCob!GASCOB_IMPORT
    End If
    
    r_rst_GasCob.Close
    Set r_rst_GasCob = Nothing
    
    ' Comparacion con fecha de proceso
    r_int_DiaAtrPro = CInt(CDate(Format(Now, "dd/mm/yyyy")) - CDate(gf_FormatoFecha(CStr(r_rst_Princi!HIPCUO_FECVCT))))
    
    ' Leyendo Gastos de Cobranzas según días de atraso -- Fecha de Proceso
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " SELECT GASCOB_IMPORT "
    g_str_Parame = g_str_Parame & "   FROM OPE_GASCOB_CAB "
    g_str_Parame = g_str_Parame & "        INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = '" & p_NumOpe & "'"
    g_str_Parame = g_str_Parame & "                             AND GASCOBCAB_FECINI <= HIPMAE_FECDES "
    g_str_Parame = g_str_Parame & "                             AND GASCOBCAB_FECFIN >= HIPMAE_FECDES "
    g_str_Parame = g_str_Parame & "        INNER JOIN OPE_GASCOB ON GASCOBCAB_CODRAN = GASCOB_CODRAN  "
    g_str_Parame = g_str_Parame & "                         AND GASCOB_DIAINI <= " & CStr(r_int_DiaAtrPro) & " "
    g_str_Parame = g_str_Parame & "                         AND GASCOB_DIAFIN >= " & CStr(r_int_DiaAtrPro) & " "
    g_str_Parame = g_str_Parame & "  WHERE GASCOBCAB_CODPRD = '" & p_CodPrd & "' "
    g_str_Parame = g_str_Parame & "    AND GASCOBCAB_CODSUB = '" & p_SubPrd & "' "
    g_str_Parame = g_str_Parame & "    AND GASCOBCAB_CODRAN <> 0 "
    
    If Not gf_EjecutaSQL(g_str_Parame, r_rst_GasCob, 3) Then
       MsgBox "Error al leer OPE_GASCOB (Operación) : " & p_NumOpe, vbInformation, modgen_g_str_NomPlt
       r_rst_GasCob.Close
       Set r_rst_GasCob = Nothing
       Exit Sub
    End If
    
    If Not (r_rst_GasCob.BOF And r_rst_GasCob.EOF) Then
       r_dbl_CobrzaPro = r_rst_GasCob!GASCOB_IMPORT
    End If
    
    r_rst_GasCob.Close
    Set r_rst_GasCob = Nothing
    
    r_dbl_Cobrza = r_dbl_CobrzaPro - r_dbl_CobrzaPag
    
    '**************************************************
    ' Total hasta la fecha del proceso
    r_dbl_Capita = r_rst_Princi!HIPCUO_CAPITA - r_rst_Princi!HIPCUO_CAPPAG
    r_dbl_Intere = r_rst_Princi!HIPCUO_INTERE - r_rst_Princi!HIPCUO_INTPAG
    r_dbl_SegDes = r_rst_Princi!HIPCUO_DESORG - r_rst_Princi!HIPCUO_DESPAG
    r_dbl_SegInm = r_rst_Princi!HIPCUO_VIVORG - r_rst_Princi!HIPCUO_VIVPAG
    r_dbl_Portes = r_rst_Princi!HIPCUO_OTRORG - r_rst_Princi!HIPCUO_OTRPAG
    r_dbl_CapBBP = r_rst_Princi!HIPCUO_CAPBBP - r_rst_Princi!HIPCUO_CBPPAG
    r_dbl_IntBBP = r_rst_Princi!HIPCUO_INTBBP - r_rst_Princi!HIPCUO_IBPPAG
    r_dbl_IntCom = r_rst_Princi!HIPCUO_INTCOM - r_rst_Princi!HIPCUO_ICOPAG
    r_dbl_IntMor = r_rst_Princi!HIPCUO_INTMOR - r_rst_Princi!HIPCUO_IMOPAG
    r_dbl_GasCob = r_rst_Princi!HIPCUO_GASCOB - r_rst_Princi!HIPCUO_GCOPAG
    r_dbl_OtrGas = r_rst_Princi!HIPCUO_OTRGAS - r_rst_Princi!HIPCUO_OTGPAG
    r_dbl_TotCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_SegDes + r_dbl_SegInm + r_dbl_Portes + r_dbl_CapBBP + r_dbl_IntBBP + r_dbl_IntCom + r_dbl_IntMor + r_dbl_GasCob + r_dbl_OtrGas
    r_dbl_ImpITF = gf_NueImp_Numero(gf_Truncar_Numero(r_dbl_TotCuo * (p_PorITF / 100), 2))
    r_dbl_TotPro = CDbl(Format(r_dbl_TotCuo + r_dbl_ImpITF, "###0.00"))
                
    r_rst_Princi.Close
    Set r_rst_Princi = Nothing
    
    ' VERIFICA SI LA DIFERENCIA ES EL INTERES POR DIAS DE ATRASO SEGUN LA FECHA DE PAGO QUE SE REALIZO EN EL BANCO
    If Abs(Format(CDbl(p_TotDep) - CDbl(r_dbl_TotPro), "##0.00")) = r_dbl_Cobrza And r_dbl_Cobrza > 0 Then
        'Actualizando en CRE_HIPCUO
        g_str_Parame = "UPDATE CRE_HIPCUO SET "
        g_str_Parame = g_str_Parame & "HIPCUO_GASCOB = HIPCUO_GASCOB - " & Format(r_dbl_Cobrza, "#######0.00") & " "
        g_str_Parame = g_str_Parame & "WHERE "
        g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
        g_str_Parame = g_str_Parame & "HIPCUO_NUMCUO = " & CStr(p_NumCuo) & " AND "
        g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
        g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 "
    
        If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 2) Then
           MsgBox "Error al actualzar en CRE_HIPCUO (Operación) : " & p_NumOpe, vbInformation, modgen_g_str_NomPlt
        End If
    End If
    
End Sub

Private Sub fs_Graba_ReportPA(ByVal p_NumOpe As String, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_FecPag As String, ByVal p_Import As Double, ByVal p_TipPag As Integer, ByVal p_NumCuo As Integer, ByVal p_DesErr As String)
   'Grabando en DAO
   moddat_g_str_CadDAO = "SELECT * FROM RPT_COBBCO WHERE COBBCO_NUMOPE = ' '"
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
   moddat_g_rst_RecDAO("COBBCO_NUMOPE") = p_NumOpe
   moddat_g_rst_RecDAO("COBBCO_DOCIDE") = CStr(p_TipDoc) & "-" & Trim(p_NumDoc & "")
   moddat_g_rst_RecDAO("COBBCO_NOMCLI") = moddat_gf_Buscar_NomCliPA(CStr(p_TipDoc), Trim(p_NumDoc))
   moddat_g_rst_RecDAO("COBBCO_FECPAG") = gf_FormatoFecha(p_FecPag)
   moddat_g_rst_RecDAO("COBBCO_IMPORT") = p_Import
   moddat_g_rst_RecDAO("COBBCO_TIPPAG") = moddat_gf_Consulta_ParDes("240", CStr(p_TipPag))
   moddat_g_rst_RecDAO("COBBCO_NUMCUO") = p_NumCuo
   moddat_g_rst_RecDAO("COBBCO_DESERR") = p_DesErr
   moddat_g_rst_RecDAO.Update
   DoEvents
   
   moddat_g_rst_RecDAO.Close
   DoEvents
End Sub

Private Function moddat_gf_Buscar_NomCliPA(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As String
   moddat_gf_Buscar_NomCliPA = ""
   
   g_str_Parame = "SELECT * FROM CRE_AHOCLI WHERE "
   g_str_Parame = g_str_Parame & "AHOCLI_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "AHOCLI_NUMDOC = '" & Trim(p_NumDoc) & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Buscar_NomCliPA = Trim(g_rst_Listas!AHOCLI_APEPAT) & " " & Trim(g_rst_Listas!AHOCLI_APEMAT) & " " & Trim(g_rst_Listas!AHOCLI_NOMBRE)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

