VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Desemb_06_ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2730
   ClientLeft      =   5220
   ClientTop       =   4575
   ClientWidth     =   10890
   Icon            =   "OpeTra_frm_035.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2745
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10890
      _Version        =   65536
      _ExtentX        =   19209
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
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
            Height          =   315
            Left            =   690
            TabIndex        =   15
            Top             =   30
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios"
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
         Begin Threed.SSPanel SSPanel68 
            Height          =   315
            Left            =   690
            TabIndex        =   16
            Top             =   330
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Desembolso - Impresión de Formatos"
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
            Left            =   10380
            Top             =   60
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
            Left            =   60
            Picture         =   "OpeTra_frm_035.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   5
         Top             =   1440
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1440
            TabIndex        =   6
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   8070
            TabIndex        =   7
            Top             =   60
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
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
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   8
            Top             =   390
            Width           =   9315
            _Version        =   65536
            _ExtentX        =   16431
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
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   6630
            TabIndex        =   11
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   12
         Top             =   750
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
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
         Begin VB.CommandButton EXCEL 
            Caption         =   "EXCEL"
            Height          =   495
            Left            =   1080
            TabIndex        =   17
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10200
            Picture         =   "OpeTra_frm_035.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_035.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Imprimir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   13
         Top             =   2250
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
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
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   9315
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo de Reporte:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "frm_Desemb_06_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_dbl_ImpTas     As Double
Dim l_dbl_ImpNot     As Double
Dim l_dbl_ImpEst     As Double
Dim l_dbl_ImpEva     As Double
Dim l_dbl_ImpAdm     As Double
Dim l_dbl_ImpRed     As Double
Dim l_dbl_ImpBlq     As Double

Private Sub cmd_Imprim_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Reporte a Imprimir.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de Imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If

   moddat_g_int_TipRep = cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
   
   If moddat_g_int_TipRep > 0 Then
      Screen.MousePointer = 11
      Select Case moddat_g_int_TipRep
         Case 1:     Call fs_LiqDes
         Case 2:     Call fs_HojRes
         Case 3:     Call fs_Cronog_MiCasita
         Case 4:     Call fs_Cronog_Mivivienda_NCoCli
         Case 5:     Call fs_Cronog_Mivivienda_ConCli
         Case 7, 9:  Call fs_Cronog_Mivivienda_ConMVi
         Case 8:     Call fs_Cronog_Mivivienda_NCoMVi
         Case 10:    Call fs_Cronog_Mivivienda_NCoCof
         Case 11:    Call fs_ComPag
         Case 12
            If moddat_g_int_InsAct = 0 Then
               MsgBox "No se puede emitir este formato, porque no se realizón ninguna operación.", vbInformation, modgen_g_str_NomPlt
               Screen.MousePointer = 0
               Exit Sub
            End If
            If moddat_g_int_InsAct = moddat_g_int_TipMon Then
               MsgBox "No se puede emitir este formato, porque la Moneda de Compra-Venta es igual a la Moneda de Préstamo.", vbInformation, modgen_g_str_NomPlt
               Screen.MousePointer = 0
               Exit Sub
            End If
            Call fs_LiqTipoCambio
      End Select
      Screen.MousePointer = 0
   End If

End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub EXCEL_Click()

MsgBox ("aqui")

End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_GasAdm
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_TipRep.Clear
   
   cmb_TipRep.AddItem "LIQUIDACION DE DESEMBOLSO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   cmb_TipRep.AddItem "HOJA RESUMEN DE CREDITO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2

   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then        '"002" "011"
      cmb_TipRep.AddItem "CRONOGRAMA DE PAGOS"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 3
   ElseIf InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then    '"019" "021" "022" "023"
      cmb_TipRep.AddItem "CRONOGRAMA DE PAGOS (CLIENTE)"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 3
      
      cmb_TipRep.AddItem "CRONOGRAMA DE PAGOS (COFIDE)"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 8
   Else
      cmb_TipRep.AddItem "CRONOGRAMA DE PAGOS (CLIENTE TRAMO NO CONCESIONAL)"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 4

      cmb_TipRep.AddItem "CRONOGRAMA DE PAGOS (CLIENTE TRAMO CONCESIONAL)"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 5
      
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Then      '"001"
         cmb_TipRep.AddItem "CRONOGRAMA DE PAGOS (MIVIVIENDA TRAMO CONCESIONAL)"
         cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 7
      ElseIf InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Then
         cmb_TipRep.AddItem "CRONOGRAMA DE PAGOS (COFIDE TRAMO NO CONCESIONAL)"
         cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 8
      
         cmb_TipRep.AddItem "CRONOGRAMA DE PAGOS (COFIDE TRAMO CONCESIONAL)"
         cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 9
      ElseIf InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then  '"003"
         cmb_TipRep.AddItem "CRONOGRAMA DE PAGOS (MIVIVIENDA TRAMO CONCESIONAL)"
         cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 7
         
         cmb_TipRep.AddItem "CRONOGRAMA DE PAGOS (COFIDE)"
         cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 10
      End If
   End If
   
   cmb_TipRep.AddItem "COMPROBANTE DE PAGO DESEMBOLSO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 11
   cmb_TipRep.AddItem "LIQUIDACION DE TIPO DE CAMBIO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 12
   cmb_TipRep.ListIndex = -1
End Sub


Private Sub fs_LiqDes()
Dim r_rst_Genera     As ADODB.Recordset
Dim r_str_Direcc     As String
Dim r_str_Distri     As String
Dim r_str_Modali     As String

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_FORDES "
   g_str_Parame = g_str_Parame & " WHERE FORDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de Créditos
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Leyendo Tabla de Desembolso
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPDES "
   g_str_Parame = g_str_Parame & " WHERE HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Para obtener Modalidad
   r_str_Modali = ""
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
      r_str_Modali = moddat_g_arr_Genera(1).Genera_Nombre
   End If
   
   'Para obtener Dirección de Inmueble
   Call moddat_gs_Consulta_DatInm(g_rst_Princi!hipmae_numsol, r_str_Direcc, r_str_Distri)
   
   'Insertando Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "INSERT INTO RPT_FORDES("
   g_str_Parame = g_str_Parame & "FORDES_NUMOPE, "
   g_str_Parame = g_str_Parame & "FORDES_NUMSOL, "
   g_str_Parame = g_str_Parame & "FORDES_MODALI, "
   g_str_Parame = g_str_Parame & "FORDES_DIRINM, "
   g_str_Parame = g_str_Parame & "FORDES_DSTINM, "
   g_str_Parame = g_str_Parame & "FORDES_EMPSEG, "
   g_str_Parame = g_str_Parame & "FORDES_TIPSEG, "
   g_str_Parame = g_str_Parame & "FORDES_BANDES, "
   g_str_Parame = g_str_Parame & "FORDES_NUMCTA) "
   
   g_str_Parame = g_str_Parame & "VALUES ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Modali & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Direcc & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Distri & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "") & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("516", r_rst_Genera!HIPDES_BANCGO & "") & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(r_rst_Genera!HIPDES_CTACGO & "") & "' )"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(1) = "CRE_HIPDES"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_FORDES"
   crp_Imprim.DataFiles(4) = "CRE_PRODUC"
   crp_Imprim.DataFiles(5) = ""
   crp_Imprim.DataFiles(6) = ""
   
   crp_Imprim.SelectionFormula = "{RPT_FORDES.FORDES_NUMOPE} = '" & moddat_g_str_NumOpe & "' "
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_LIQDES_11.RPT"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_HojRes()
Dim r_rst_Genera     As ADODB.Recordset
Dim r_str_Direcc     As String
Dim r_str_Distri     As String
Dim r_str_Modali     As String
Dim r_dbl_PorITF     As Double
Dim r_dbl_IntMor     As Double
Dim r_dbl_PrePag     As Double
Dim r_dbl_LevHip     As Double
Dim r_dbl_CamTas     As Double
Dim r_dbl_CobJud     As Double
Dim r_dbl_CanMVi     As Double
Dim r_dbl_CobDi1     As Double
Dim r_dbl_CobIm1     As Double
Dim r_dbl_CobDi2     As Double
Dim r_dbl_CobIm2     As Double
Dim r_dbl_CobDi3     As Double
Dim r_dbl_CobIm3     As Double
Dim r_int_Indice     As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_FORDES "
   g_str_Parame = g_str_Parame & " WHERE FORDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de Créditos
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Leyendo Tabla de Desembolso
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPDES "
   g_str_Parame = g_str_Parame & " WHERE HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Para obtener Modalidad
   r_str_Modali = ""
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
      r_str_Modali = moddat_g_arr_Genera(1).Genera_Nombre
   End If
   
   'Para obtener Interes Moratorio
   r_dbl_IntMor = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "201") Then
      r_dbl_IntMor = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   'Para obtener ITF
   r_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
   
   'Para obtener Dirección de Inmueble
   Call moddat_gs_Consulta_DatInm(g_rst_Princi!hipmae_numsol, r_str_Direcc, r_str_Distri)
   
   'Otras Comisiones - Prepagos
   r_dbl_PrePag = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "501") Then
      r_dbl_PrePag = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   'Otras Comisiones - Levantamiento de Hipoteca
   r_dbl_LevHip = 0
   'If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "701") Then
   '   r_dbl_LevHip = moddat_g_arr_Genera(1).Genera_Cantid
   'End If
   
   'Otras Comisiones - Cambio de Fecha, Tasa de Interes, Moneda o Cuota
   r_dbl_CamTas = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "702") Then
      r_dbl_CamTas = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   'Otras Comisiones - Cobranza Judicial
   r_dbl_CobJud = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "703") Then
      r_dbl_CobJud = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   'Otras Comisiones - Caducidad del Servicio MiVivienda
   r_dbl_CanMVi = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "901") Then
      r_dbl_CanMVi = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   'Gastos de Cobranzas
   r_dbl_CobDi1 = 0
   r_dbl_CobIm1 = 0
   r_dbl_CobDi2 = 0
   r_dbl_CobIm2 = 0
   r_dbl_CobDi3 = 0
   r_dbl_CobIm3 = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM OPE_GASCOB "
   g_str_Parame = g_str_Parame & " WHERE GASCOB_CODPRD = '" & g_rst_Princi!HIPMAE_CODPRD & "' "
   g_str_Parame = g_str_Parame & "   AND GASCOB_CODSUB = '" & g_rst_Princi!HIPMAE_CODSUB & "' "
   g_str_Parame = g_str_Parame & "   AND GASCOB_IMPORT > 0 "
   g_str_Parame = g_str_Parame & " ORDER BY GASCOB_DIAINI ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      
      r_int_Indice = 1
      Do While Not g_rst_Genera.EOF
         If r_int_Indice = 1 Then
            r_dbl_CobDi1 = g_rst_Genera!GasCob_DiaIni
            r_dbl_CobIm1 = g_rst_Genera!GasCob_Import
         ElseIf r_int_Indice = 2 Then
            r_dbl_CobDi2 = g_rst_Genera!GasCob_DiaIni
            r_dbl_CobIm2 = g_rst_Genera!GasCob_Import
         ElseIf r_int_Indice = 3 Then
            r_dbl_CobDi3 = g_rst_Genera!GasCob_DiaIni
            r_dbl_CobIm3 = g_rst_Genera!GasCob_Import
         End If
         
         r_int_Indice = r_int_Indice + 1
         g_rst_Genera.MoveNext
         DoEvents
      Loop

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   'Insertando Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "INSERT INTO RPT_FORDES("
   g_str_Parame = g_str_Parame & "FORDES_NUMOPE, "
   g_str_Parame = g_str_Parame & "FORDES_NUMSOL, "
   g_str_Parame = g_str_Parame & "FORDES_MODALI, "
   g_str_Parame = g_str_Parame & "FORDES_DIRINM, "
   g_str_Parame = g_str_Parame & "FORDES_DSTINM, "
   g_str_Parame = g_str_Parame & "FORDES_EMPSEG, "
   g_str_Parame = g_str_Parame & "FORDES_TIPSEG, "
   g_str_Parame = g_str_Parame & "FORDES_BANDES, "
   g_str_Parame = g_str_Parame & "FORDES_NUMCTA,"
   g_str_Parame = g_str_Parame & "FORDES_TASMOR, "
   g_str_Parame = g_str_Parame & "FORDES_PORITF, "
   g_str_Parame = g_str_Parame & "FORDES_GASTAS, "
   g_str_Parame = g_str_Parame & "FORDES_GASNOT, "
   g_str_Parame = g_str_Parame & "FORDES_ESTTIT, "
   g_str_Parame = g_str_Parame & "FORDES_EVACRE, "
   g_str_Parame = g_str_Parame & "FORDES_ADMTAS, "
   g_str_Parame = g_str_Parame & "FORDES_REDCON, "
   g_str_Parame = g_str_Parame & "FORDES_BLQREG, "
   g_str_Parame = g_str_Parame & "FORDES_PREPAG, "
   g_str_Parame = g_str_Parame & "FORDES_LEVHIP, "
   g_str_Parame = g_str_Parame & "FORDES_CAMTAS, "
   g_str_Parame = g_str_Parame & "FORDES_COBJUD, "
   g_str_Parame = g_str_Parame & "FORDES_COBDI1, "
   g_str_Parame = g_str_Parame & "FORDES_COBDI2, "
   g_str_Parame = g_str_Parame & "FORDES_COBDI3, "
   g_str_Parame = g_str_Parame & "FORDES_COBIM1, "
   g_str_Parame = g_str_Parame & "FORDES_COBIM2, "
   g_str_Parame = g_str_Parame & "FORDES_COBIM3) "
   
   g_str_Parame = g_str_Parame & "VALUES ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Modali & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Direcc & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Distri & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "") & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("516", r_rst_Genera!HIPDES_BANCGO & "") & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(r_rst_Genera!HIPDES_CTACGO & "") & "', "
   g_str_Parame = g_str_Parame & CStr(r_dbl_IntMor) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_PorITF) & ", "
   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpTas) & ", "
   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpNot) & ", "
   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpEst) & ", "
   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpEva) & ", "
   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpAdm) & ", "
   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpRed) & ", "
   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpBlq) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_PrePag) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_LevHip) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_CamTas) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_CobJud) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_CobDi1) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_CobDi2) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_CobDi3) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_CobIm1) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_CobIm2) & ", "
   g_str_Parame = g_str_Parame & CStr(r_dbl_CobIm3) & ") "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(1) = "CRE_HIPDES"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_FORDES"
   crp_Imprim.DataFiles(4) = "CRE_PRODUC"
   crp_Imprim.DataFiles(5) = "TRA_EVALEG"
   crp_Imprim.DataFiles(6) = "TRA_POLIZA"
   crp_Imprim.SelectionFormula = "{RPT_FORDES.FORDES_NUMOPE} = '" & moddat_g_str_NumOpe & "' "
   
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "002" Or moddat_g_str_CodPrd = "011" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_11.RPT"
   ElseIf InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_12.RPT"
   ElseIf InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_13.RPT"
   ElseIf InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_14.RPT"
   End If
   
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_MiCasita()
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   crp_Imprim.DataFiles(4) = ""
   crp_Imprim.DataFiles(5) = ""
   crp_Imprim.DataFiles(6) = ""
   
   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 1 "
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_11.RPT"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_Mivivienda_NCoCli()
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   crp_Imprim.DataFiles(4) = ""
   crp_Imprim.DataFiles(5) = ""
   crp_Imprim.DataFiles(6) = ""
   
   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 1 "
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_12.RPT"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_Mivivienda_ConCli()
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   crp_Imprim.DataFiles(4) = ""
   crp_Imprim.DataFiles(5) = ""
   crp_Imprim.DataFiles(6) = ""
   
   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 2 "
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_13.RPT"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_Mivivienda_ConMVi()
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   
   If InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then      '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018"
      crp_Imprim.DataFiles(4) = "TRA_EVACOF"
   Else
      crp_Imprim.DataFiles(4) = ""
   End If
   
   crp_Imprim.DataFiles(5) = ""
   crp_Imprim.DataFiles(6) = ""
   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 4 "
   
   If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then        '"001" "003"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_14.RPT"
   ElseIf InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then  '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_15.RPT"
   End If
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_Mivivienda_NCoMVi()
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   
   If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then  '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
      crp_Imprim.DataFiles(4) = "TRA_EVACOF"
   Else
      crp_Imprim.DataFiles(4) = ""
   End If
   
   crp_Imprim.DataFiles(5) = ""
   crp_Imprim.DataFiles(6) = ""
   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 3 "
   
   If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_16.RPT"
   End If
   
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_Mivivienda_NCoCof()
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   crp_Imprim.DataFiles(4) = "TRA_EVACOF"
   crp_Imprim.DataFiles(5) = ""
   crp_Imprim.DataFiles(6) = ""
   
   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 5 "
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_17.RPT"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_ComPag()
Dim r_str_CodSuc     As String
Dim r_str_NumMov     As String
Dim r_str_FecMov     As String

   'Buscando en OPE_CAJMOV
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM OPE_CAJMOV "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_TIPMOV = 1103 "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   r_str_CodSuc = Trim(g_rst_Princi!CAJMOV_SUCMOV)
   r_str_NumMov = CStr(g_rst_Princi!CAJMOV_NUMMOV)
   r_str_FecMov = CStr(g_rst_Princi!CAJMOV_FECMOV)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Borrar Spool de PC (Cabecera)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_COMPGC "
   g_str_Parame = g_str_Parame & " WHERE COMPGC_CODTER = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   'Borrar Spool de PC (Detalle)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_COMPGD "
   g_str_Parame = g_str_Parame & " WHERE COMPGD_CODTER = '" & modgen_g_str_NombPC & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call opecaj_gs_ComPago(r_str_CodSuc, r_str_NumMov, r_str_FecMov, 1, 1)
   Screen.MousePointer = 0
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "RPT_COMPGC"
   crp_Imprim.DataFiles(1) = "RPT_COMPGD"
   crp_Imprim.DataFiles(2) = ""
   crp_Imprim.DataFiles(3) = ""
   crp_Imprim.DataFiles(4) = ""
   crp_Imprim.DataFiles(5) = ""
   crp_Imprim.DataFiles(6) = ""

   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_COMPAG_01.RPT"
   crp_Imprim.SelectionFormula = "{RPT_COMPGC.COMPGC_CODTER} = '" & modgen_g_str_NombPC & "'"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_LiqTipoCambio()
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPDES"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = ""
   crp_Imprim.DataFiles(4) = ""
   crp_Imprim.DataFiles(5) = ""
   crp_Imprim.DataFiles(6) = ""

   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_LIQDES_12.RPT"
   crp_Imprim.SelectionFormula = "{CRE_HIPDES.HIPDES_NUMOPE} = '" & moddat_g_str_NumOpe & "'"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GasAdm()
   'Inicializando Variables para Hoja Resumen
   l_dbl_ImpTas = 0
   l_dbl_ImpNot = 0
   l_dbl_ImpEst = 0
   l_dbl_ImpEva = 0
   l_dbl_ImpAdm = 0
   l_dbl_ImpRed = 0
   l_dbl_ImpBlq = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_GASADM "
   g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         Select Case g_rst_Princi!GASADM_CODGAS
            Case 11: l_dbl_ImpTas = g_rst_Princi!GASADM_IMPORT
            Case 12: l_dbl_ImpNot = g_rst_Princi!GASADM_IMPORT
            Case 14: l_dbl_ImpEst = g_rst_Princi!GASADM_IMPORT
            Case 15, 22: l_dbl_ImpEva = g_rst_Princi!GASADM_IMPORT
            Case 16: l_dbl_ImpBlq = g_rst_Princi!GASADM_IMPORT
            Case 20: l_dbl_ImpAdm = g_rst_Princi!GASADM_IMPORT
            Case 21: l_dbl_ImpRed = g_rst_Princi!GASADM_IMPORT
         End Select
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

