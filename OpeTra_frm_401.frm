VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Desemb_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   5070
   ClientLeft      =   5220
   ClientTop       =   4575
   ClientWidth     =   10890
   Icon            =   "OpeTra_frm_401.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5085
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10890
      _Version        =   65536
      _ExtentX        =   19209
      _ExtentY        =   8969
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
         TabIndex        =   3
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
            TabIndex        =   14
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
            TabIndex        =   15
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
            Picture         =   "OpeTra_frm_401.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   795
         Left            =   30
         TabIndex        =   4
         Top             =   1440
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1440
            TabIndex        =   5
            Top             =   60
            Width           =   5955
            _Version        =   65536
            _ExtentX        =   10504
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
            Left            =   8970
            TabIndex        =   6
            Top             =   60
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
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
            TabIndex        =   7
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
            Left            =   7710
            TabIndex        =   10
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   90
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   420
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   11
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
         Begin VB.CommandButton cmd_resumen 
            Caption         =   "exc"
            Height          =   585
            Left            =   750
            TabIndex        =   18
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10200
            Picture         =   "OpeTra_frm_401.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_401.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Imprimir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   2775
         Left            =   30
         TabIndex        =   12
         Top             =   2250
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
         _ExtentY        =   4895
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
         Begin VB.CheckBox chk_VistaPrevia 
            Caption         =   "Solo Vista Preliminar"
            Height          =   315
            Left            =   1470
            TabIndex        =   17
            Top             =   2400
            Value           =   1  'Checked
            Width           =   1755
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listado 
            Height          =   2235
            Left            =   1410
            TabIndex        =   16
            Top             =   90
            Width           =   9345
            _ExtentX        =   16484
            _ExtentY        =   3942
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo de Reporte:"
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   150
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "frm_Desemb_06"
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

Private Sub cmd_ExpExc_Click()

      
End Sub

Private Sub cmd_resumen_Click()
  '********* 22102020   inicio rat
         Dim r_str_variable As String
         r_str_variable = ""
         r_str_variable = r_str_variable & "'0221900482'"
          g_str_Parame = ""
          g_str_Parame = g_str_Parame & "select hipmae_numope as operacion, trim(hipmae_ndocli) as documento from cre_hipmae"
          g_str_Parame = g_str_Parame & "  where hipmae_numope in (" & r_str_variable & ") "
            MsgBox (g_str_Parame)
       '********* 22102020   fin rat
          If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
               Exit Sub
          End If
          
          Dim r_int_ConVer As Integer
            g_rst_Princi.MoveFirst
            r_int_ConVer = 1
               Do While Not g_rst_Princi.EOF
                   fs_GenExcNuevo2 (g_rst_Princi!OPERACION)
                   r_int_ConVer = r_int_ConVer + 1
                   g_rst_Princi.MoveNext
               DoEvents
                Loop
               g_rst_Princi.Close
               Set g_rst_Princi = Nothing
        '********* 22102020   fin rat
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
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
   grd_Listado.Clear
   grd_Listado.FixedCols = 0
   grd_Listado.Rows = 1
   grd_Listado.ColWidth(0) = 0
   grd_Listado.ColWidth(1) = 8000
   grd_Listado.ColWidth(2) = 1000
   
   grd_Listado.TextMatrix(0, 0) = "ID"
   grd_Listado.TextMatrix(0, 1) = "Nombre del Reporte"
   grd_Listado.TextMatrix(0, 2) = "# Copias"
   
'   grd_Listado.AddItem "1" & vbTab & "LIQUIDACION DE DESEMBOLSO" & vbTab & "3"
grd_Listado.AddItem "1" & vbTab & "LIQUIDACION DE DESEMBOLSO" & vbTab & "0"
'   grd_Listado.AddItem "2" & vbTab & "HOJA RESUMEN DE CREDITO" & vbTab & "2"
grd_Listado.AddItem "2" & vbTab & "HOJA RESUMEN DE CREDITO" & vbTab & "1"
   
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "002" Or moddat_g_str_CodPrd = "011" Then
'      grd_Listado.AddItem "3" & vbTab & "CRONOGRAMA DE PAGOS" & vbTab & "2"
      grd_Listado.AddItem "3" & vbTab & "CRONOGRAMA DE PAGOS" & vbTab & "0"
   ElseIf InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
     grd_Listado.AddItem "3" & vbTab & "CRONOGRAMA DE PAGOS (CLIENTE)" & vbTab & "0"
'       grd_Listado.AddItem "3" & vbTab & "CRONOGRAMA DE PAGOS (CLIENTE)" & vbTab & "2"
            
'      grd_Listado.AddItem "8" & vbTab & "CRONOGRAMA DE PAGOS (COFIDE)" & vbTab & "1"
       grd_Listado.AddItem "8" & vbTab & "CRONOGRAMA DE PAGOS (COFIDE)" & vbTab & "0"
   Else
      grd_Listado.AddItem "4" & vbTab & "CRONOGRAMA DE PAGOS (CLIENTE TRAMO NO CONCESIONAL)" & vbTab & "2"
      
      grd_Listado.AddItem "5" & vbTab & "CRONOGRAMA DE PAGOS (CLIENTE TRAMO CONCESIONAL)" & vbTab & "2"
            
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Then
         grd_Listado.AddItem "7" & vbTab & "CRONOGRAMA DE PAGOS (MIVIVIENDA TRAMO CONCESIONAL)" & vbTab & "1"
         
      ElseIf InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Then
         grd_Listado.AddItem "8" & vbTab & "CRONOGRAMA DE PAGOS (COFIDE TRAMO NO CONCESIONAL)" & vbTab & "1"
               
         grd_Listado.AddItem "9" & vbTab & "CRONOGRAMA DE PAGOS (COFIDE TRAMO CONCESIONAL)" & vbTab & "1"
         
      ElseIf InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
         grd_Listado.AddItem "7" & vbTab & "CRONOGRAMA DE PAGOS (MIVIVIENDA TRAMO CONCESIONAL)" & vbTab & "1"
                  
         grd_Listado.AddItem "10" & vbTab & "CRONOGRAMA DE PAGOS (COFIDE)" & vbTab & "1"
      End If
   End If
   
   grd_Listado.AddItem "11" & vbTab & "COMPROBANTE DE PAGO DESEMBOLSO" & vbTab & "2"
      
   If moddat_g_int_TipMon = 1 Then
      grd_Listado.AddItem "12" & vbTab & "LIQUIDACION DE TIPO DE CAMBIO" & vbTab & "0"
   Else
      grd_Listado.AddItem "12" & vbTab & "LIQUIDACION DE TIPO DE CAMBIO" & vbTab & "2"
   End If
   grd_Listado.AddItem "13" & vbTab & "NOTA DE ABONO" & vbTab & "0"
End Sub

Private Sub fs_LiqDes(p_Fila As Integer)
Dim r_rst_Genera     As ADODB.Recordset
Dim r_str_Direcc     As String
Dim r_str_Distri     As String
Dim r_str_Modali     As String
Dim r_int_Cont       As Integer
Dim r_rst_Princi     As ADODB.Recordset
Dim r_str_Parame     As String

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
   
   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
   If p_Fila = 13 Then
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT MN.CTABAN_CODBCO, MN.CTABAN_NUMCTA, MN.CTABAN_NUMCCI, "
      r_str_Parame = r_str_Parame & "        (SELECT TRIM(C.PARDES_DESCRI) FROM MNT_PARDES C WHERE C.PARDES_CODGRP = 513 AND C.PARDES_CODITE = MN.CTABAN_CODBCO) AS BANCO "
      r_str_Parame = r_str_Parame & "   FROM PRY_CTABAN MN "
      r_str_Parame = r_str_Parame & "  WHERE TRIM(MN.CTABAN_CODPRY) = (SELECT A.SOLINM_PRYCOD "
      r_str_Parame = r_str_Parame & "                                    FROM CRE_SOLINM A "
      r_str_Parame = r_str_Parame & "                                   WHERE A.SOLINM_NUMSOL = (SELECT B.HIPMAE_NUMSOL "
      r_str_Parame = r_str_Parame & "                                                              FROM CRE_HIPMAE B "
      r_str_Parame = r_str_Parame & "                                                             WHERE B.HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "')) "
      r_str_Parame = r_str_Parame & "    AND MN.CTABAN_SITUAC = 1 "
      r_str_Parame = r_str_Parame & "    AND MN.CTABAN_TIPMON = " & g_rst_Princi!HIPMAE_MONEDA
   
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
         Exit Sub
      End If
         
      If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
         r_rst_Princi.MoveFirst
         crp_Imprim.ParameterFields(0) = "p_nomban;" & r_rst_Princi!BANCO & ";True"
         crp_Imprim.ParameterFields(1) = "p_ctacci;" & r_rst_Princi!CTABAN_NUMCCI & ";True"
      Else
         crp_Imprim.ParameterFields(0) = "p_nomban;" & "" & ";True"
         crp_Imprim.ParameterFields(1) = "p_ctacci;" & "" & ";True"
      End If
                              
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT NVL(CASE WHEN H.SOLINM_TIPDOC_CON = 7 THEN TRIM(L.DATGEN_RAZSOC) ELSE TRIM(H.SOLINM_RAZSOC_CON) END, '-') AS NOM_CONSTRUCTOR "
      r_str_Parame = r_str_Parame & "   FROM CRE_SOLINM H "
      r_str_Parame = r_str_Parame & "   LEFT JOIN EMP_DATGEN L ON L.DATGEN_EMPTDO = H.SOLINM_TIPDOC_CON AND L.DATGEN_EMPNDO = H.SOLINM_NUMDOC_CON "
      r_str_Parame = r_str_Parame & "  WHERE H.SOLINM_NUMSOL = (SELECT B.HIPMAE_NUMSOL "
      r_str_Parame = r_str_Parame & "                             FROM CRE_HIPMAE B "
      r_str_Parame = r_str_Parame & "                            WHERE B.HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "') "
      
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
         Exit Sub
      End If
      If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
         crp_Imprim.ParameterFields(2) = "p_constr;" & r_rst_Princi!NOM_CONSTRUCTOR & ";True"
      Else
         crp_Imprim.ParameterFields(2) = "p_constr;" & "" & ";True"
      End If
   End If
   
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
            
            
   If p_Fila = 1 Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_LIQDES_11.RPT"
   ElseIf p_Fila = 13 Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_NOTABO_11.RPT"
   End If
   
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
   If chk_VistaPrevia.Value = 1 Then Exit Sub
   
   r_int_Cont = 0
   Do While r_int_Cont < Val(grd_Listado.TextMatrix(1, 2))
      crp_Imprim.Destination = crptToPrinter
      crp_Imprim.PrintReport
      r_int_Cont = r_int_Cont + 1
   Loop
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
Dim r_int_Cont       As Integer

   moddat_g_str_NumOpe = "0071100063"

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
   
   r_rst_Genera.MoveFirst
   
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
   g_str_Parame = g_str_Parame & " SELECT GASCOB_DIAINI, GASCOB_IMPORT "
   g_str_Parame = g_str_Parame & "   FROM OPE_GASCOB_CAB "
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND GASCOBCAB_FECINI <= HIPMAE_FECDES AND GASCOBCAB_FECFIN >= HIPMAE_FECDES "
   g_str_Parame = g_str_Parame & "        INNER JOIN OPE_GASCOB ON GASCOBCAB_CODRAN = GASCOB_CODRAN AND GASCOB_IMPORT > 0 "
   g_str_Parame = g_str_Parame & "  WHERE GASCOBCAB_CODPRD = '" & g_rst_Princi!HIPMAE_CODPRD & "' "
   g_str_Parame = g_str_Parame & "    AND GASCOBCAB_CODSUB = '" & g_rst_Princi!HIPMAE_CODSUB & "' "
   g_str_Parame = g_str_Parame & "    AND GASCOBCAB_CODRAN <> 0 "
   g_str_Parame = g_str_Parame & "  ORDER BY GASCOB_DIAINI ASC "
   
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
   
   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(1) = "CRE_HIPDES"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_FORDES"
   crp_Imprim.DataFiles(4) = "CRE_PRODUC"
   crp_Imprim.DataFiles(5) = "TRA_EVALEG"
   crp_Imprim.DataFiles(6) = "TRA_POLIZA"
   crp_Imprim.SelectionFormula = "{RPT_FORDES.FORDES_NUMOPE} = '" & moddat_g_str_NumOpe & "' "
   
   'If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "002" Or moddat_g_str_CodPrd = "011" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
   '   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_11.RPT"
   'ElseIf InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
   '   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_12.RPT"
   'ElseIf InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Then
   '   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_13.RPT"
   'ElseIf InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "012" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Then
   '   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_14.RPT"
   'End If
   
   Select Case moddat_g_str_CodPrd
         Case "002", "006", "011" 'MICASITA
               crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_11.RPT"
         Case "024" 'TECHO PROPIO
               crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_15.RPT" '"OPE_HOJRES_11.RPT"
         Case "003" 'CME
               crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_12.RPT"
         Case "004" 'MIHOGAR
               crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_13.RPT"
         Case Else 'MIVIVIENDA
              crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_14.RPT"
   End Select
   
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
   If chk_VistaPrevia.Value = 1 Then Exit Sub
   
   r_int_Cont = 0
   Do While r_int_Cont < Val(grd_Listado.TextMatrix(2, 2))
      crp_Imprim.Destination = crptToPrinter
      crp_Imprim.PrintReport
      r_int_Cont = r_int_Cont + 1
   Loop
End Sub

Private Sub fs_Cronog_MiCasita()
Dim r_int_Cont As Integer
   
   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
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
   If chk_VistaPrevia.Value = 1 Then Exit Sub
   
   r_int_Cont = 0
   Do While r_int_Cont < Val(grd_Listado.TextMatrix(3, 2))
      crp_Imprim.Destination = crptToPrinter
      crp_Imprim.PrintReport
      r_int_Cont = r_int_Cont + 1
   Loop
End Sub

Private Sub fs_Cronog_Mivivienda_NCoCli()
Dim r_int_Cont As Integer
   
   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
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
   
   crp_Imprim.Action = 1
   If chk_VistaPrevia.Value = 1 Then Exit Sub
   
   r_int_Cont = 0
   Do While r_int_Cont < Val(grd_Listado.TextMatrix(3, 2))
      crp_Imprim.Destination = crptToPrinter
      crp_Imprim.PrintReport
      r_int_Cont = r_int_Cont + 1
   Loop
End Sub

Private Sub fs_Cronog_Mivivienda_ConCli()
Dim r_int_Cont As Integer
   
   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
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
   
   crp_Imprim.Action = 1
   If chk_VistaPrevia.Value = 1 Then Exit Sub
   
   r_int_Cont = 0
   Do While r_int_Cont < Val(grd_Listado.TextMatrix(4, 2))
      crp_Imprim.Destination = crptToPrinter
      crp_Imprim.PrintReport
      r_int_Cont = r_int_Cont + 1
   Loop
End Sub

Private Sub fs_Cronog_Mivivienda_ConMVi()
Dim r_int_Cont As Integer

   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   
   If InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Then
      crp_Imprim.DataFiles(4) = "TRA_EVACOF"
   Else
      crp_Imprim.DataFiles(4) = ""
   End If
   
   crp_Imprim.DataFiles(5) = ""
   crp_Imprim.DataFiles(6) = ""
   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 4 "
   
   If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_14.RPT"
   ElseIf InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_15.RPT"
   End If
   
   crp_Imprim.Action = 1
   If chk_VistaPrevia.Value = 1 Then Exit Sub
   
   r_int_Cont = 0
   Do While r_int_Cont < Val(grd_Listado.TextMatrix(6, 2))
      crp_Imprim.Destination = crptToPrinter
      crp_Imprim.PrintReport
      r_int_Cont = r_int_Cont + 1
   Loop
End Sub

Private Sub fs_Cronog_Mivivienda_NCoMVi()
Dim r_int_Cont As Integer
   
   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   
   If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
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
   
   crp_Imprim.Action = 1
   If chk_VistaPrevia.Value = 1 Then Exit Sub
   
   r_int_Cont = 0
   Do While r_int_Cont < Val(grd_Listado.TextMatrix(5, 2))
      crp_Imprim.Destination = crptToPrinter
      crp_Imprim.PrintReport
      r_int_Cont = r_int_Cont + 1
   Loop
End Sub

Private Sub fs_Cronog_Mivivienda_NCoCof()
Dim r_int_Cont As Integer
   
   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
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
   
   crp_Imprim.Action = 1
   If chk_VistaPrevia.Value = 1 Then Exit Sub
   
   r_int_Cont = 0
   Do While r_int_Cont < Val(grd_Listado.TextMatrix(10, 2))
      crp_Imprim.Destination = crptToPrinter
      crp_Imprim.PrintReport
      r_int_Cont = r_int_Cont + 1
   Loop
End Sub

Private Sub fs_ComPag()
Dim r_str_CodSuc     As String
Dim r_str_NumMov     As String
Dim r_str_FecMov     As String
Dim r_int_Cont       As Integer
Dim r_int_Posicion   As Integer

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
   
   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
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
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   If chk_VistaPrevia.Value = 1 Then Exit Sub
   
   If grd_Listado.Rows = 6 Then
      r_int_Posicion = 4
   ElseIf grd_Listado.Rows = 9 Then
      r_int_Posicion = 7
   End If
   
   r_int_Cont = 0
   Do While r_int_Cont < Val(grd_Listado.TextMatrix(r_int_Posicion, 2))
      crp_Imprim.Destination = crptToPrinter
      crp_Imprim.PrintReport
      r_int_Cont = r_int_Cont + 1
   Loop
End Sub

Private Sub fs_LiqTipoCambio()
Dim r_int_Cont As Integer
Dim r_int_Posicion As Integer
   
   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
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
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   If chk_VistaPrevia.Value = 1 Then Exit Sub
   
   If grd_Listado.Rows = 6 Then
      r_int_Posicion = 5
   ElseIf grd_Listado.Rows = 9 Then
      r_int_Posicion = 8
   End If
   
   r_int_Cont = 0
   Do While r_int_Cont < Val(grd_Listado.TextMatrix(r_int_Posicion, 2))
      crp_Imprim.Destination = crptToPrinter
      crp_Imprim.PrintReport
      r_int_Cont = r_int_Cont + 1
   Loop
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

Private Sub grd_Listado_DblClick()
   Dim r_int_Cantidad As String
   
   r_int_Cantidad = InputBox("Ingrese cantidad de copia(s) a imprimir :", "Impresion", grd_Listado.TextMatrix(grd_Listado.Row, 2))
   If r_int_Cantidad <> "" Then
      grd_Listado.TextMatrix(grd_Listado.Row, 2) = Val(r_int_Cantidad)
   End If
End Sub

Private Sub grd_Listado_SelChange()
   If grd_Listado.Rows > 2 Then
      grd_Listado.RowSel = grd_Listado.Row
   End If
End Sub


Private Function fs_GenExcNuevo(var As String) As String
Dim r_rst_Princi      As ADODB.Recordset
Dim r_rst_Prindet      As ADODB.Recordset
Dim r_obj_Excel       As EXCEL.Application
Dim r_int_NumFil      As Integer
Dim r_str_Parame      As String

      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT CRE_HIPCUO.HIPCUO_NUMOPE, CRE_HIPCUO.HIPCUO_NUMCUO, HIPCUO_FECVCT, CRE_HIPCUO.HIPCUO_CAPITA, CRE_HIPCUO.HIPCUO_INTERE,"
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_DESORG, CRE_HIPCUO.HIPCUO_VIVORG, CRE_HIPCUO.HIPCUO_OTRORG, CRE_HIPCUO.HIPCUO_SALCAP, CRE_HIPMAE.HIPMAE_TDOCLI,"
      r_str_Parame = r_str_Parame & "CAST(CRE_HIPMAE.HIPMAE_NDOCLI AS VARCHAR(30)) AS HIPMAE_NDOCLI , CRE_HIPMAE.HIPMAE_PLAANO,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_CUOANO,CRE_HIPMAE.HIPMAE_PERGRA, CRE_HIPMAE.HIPMAE_NUMCUO, CRE_HIPMAE.HIPMAE_FECDES, CRE_HIPMAE.HIPMAE_MONEDA,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_MTOPRE, CRE_HIPMAE.HIPMAE_INTCAP, CRE_HIPMAE.HIPMAE_TASINT, CRE_HIPMAE.HIPMAE_COSEFE, CLI_DATGEN.DATGEN_APEPAT,"
      r_str_Parame = r_str_Parame & "CLI_DATGEN.DATGEN_APEMAT , CLI_DATGEN.DATGEN_APECAS, CLI_DATGEN.DATGEN_NOMBRE, CRE_PRODUC.PRODUC_DESCRI"
      r_str_Parame = r_str_Parame & " From "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO CRE_HIPCUO,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE CRE_HIPMAE,"
      r_str_Parame = r_str_Parame & "CLI_DATGEN CLI_DATGEN,"
      r_str_Parame = r_str_Parame & "CRE_PRODUC CRE_PRODUC"
      r_str_Parame = r_str_Parame & " Where "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_NUMOPE = CRE_HIPMAE.HIPMAE_NUMOPE AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_TDOCLI = CLI_DATGEN.DATGEN_TIPDOC AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_NDOCLI = CLI_DATGEN.DATGEN_NUMDOC AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_CODPRD = CRE_PRODUC.Produc_Codigo AND "
        r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_TIPCRO = 1 AND "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_NUMOPE = '" & var & "'"
   
   
'   HIPCUO_TIPCRO
   
'MsgBox (r_str_Parame)

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

   Set r_obj_Excel = New EXCEL.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      'MARGENES
      .PageSetup.LeftMargin = Application.CentimetersToPoints(1.5)
      .PageSetup.RightMargin = Application.CentimetersToPoints(0.4)
      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
      .Columns("A").ColumnWidth = 10
      .Columns("B").ColumnWidth = 10
       .Columns("C").ColumnWidth = 8
      .Columns("D").ColumnWidth = 8
      .Columns("E").ColumnWidth = 8
      .Columns("F").ColumnWidth = 10
      .Columns("G").ColumnWidth = 10
      .Columns("H").ColumnWidth = 8
      
      .Columns("I").ColumnWidth = 8
      .Range(.Cells(1, 1), .Cells(600, 12)).Font.Name = "Arial (Western)"
      .Range(.Cells(1, 1), .Cells(600, 12)).Font.Size = 8
      .Range(.Cells(1, 1), .Cells(600, 12)).RowHeight = 14
      
      .Pictures.Insert(g_str_RutLog & "\" & "image001.gif").Select
      
      .Cells(7, 1) = "CRONOGRAMA DE PAGOS"
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
      
      .Cells(2, 9) = "PAGOS"
          .Cells(2, 9).HorizontalAlignment = xlHAlignCenter
      .Cells(3, 9) = Format(date, "dd/mm/yyyy")
      .Cells(3, 9).HorizontalAlignment = xlHAlignCenter
      .Cells(4, 9) = Format(Time, "hh:mm:ss")
      .Cells(4, 9).HorizontalAlignment = xlHAlignCenter
      .Cells(5, 9) = "1"
      .Cells(5, 9).HorizontalAlignment = xlHAlignCenter
'      .Range(.Cells(2, 9), .Cells(5, 9)).HorizontalAlignment = xlHAlignRight
      
'      .Cells(10, 1) = "Nro Operación:"
'      .Cells(10, 2) = r_rst_Princi!HIPCUO_NUMOPE
'      .Range(.Cells(10, 1), .Cells(10, 10)).Font.Bold = True



       .Range("A10:B10").Merge
      .Range("A10") = "Nro Operación:"
       .Range("C10:E10").Merge
       .Range("C10") = r_rst_Princi!HIPCUO_NUMOPE
      
       .Range("A11:B11").Merge
      .Range("A11") = "Cliente:"
       .Range("C11:E11").Merge
       .Range("C11") = Trim(r_rst_Princi!DATGEN_APEPAT) & " " & Trim(r_rst_Princi!DATGEN_APEMAT) & " " & Trim(r_rst_Princi!DATGEN_NOMBRE)
      
       .Range("A12:B12").Merge
      .Range("A12") = "Producto:"
       .Range("C12:F12").Merge
       .Range("C12") = Trim(r_rst_Princi!PRODUC_DESCRI)
      
       .Range("A13:B13").Merge
      .Range("A13") = "Moneda:"
       .Range("C13:E13").Merge
       .Range("C13") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "SOLES", "DOLARES AMERICANOS")
       
       
      .Range("A14:B14").Merge
      .Range("A14") = "Plazo:"
       .Range("C14:E14").Merge
       .Range("C14") = Trim(r_rst_Princi!HIPMAE_PLAANO & " Años")
       
       
       .Range("A15:B15").Merge
      .Range("A15") = "Período Gracia:"
       .Range("C15:E15").Merge
       .Range("C15") = Trim(r_rst_Princi!HIPMAE_PERGRA & " Mes(es)")
       
        Dim a As String
             If r_rst_Princi!HIPMAE_CUOANO = 1 Then
               a = "NO"
             ElseIf r_rst_Princi!HIPMAE_CUOANO = 2 Then
               a = "JULIO"
             ElseIf r_rst_Princi!HIPMAE_CUOANO = 3 Then
               a = "DICIEMBRE"
             ElseIf r_rst_Princi!HIPMAE_CUOANO = 4 Then
               a = "JULIO Y DICIEMBRE"
             Else
               a = ""
             End If
       
       .Range("A16:B16").Merge
      .Range("A16") = "Cuotas Extraord:"
       .Range("C16:E16").Merge
       .Range("C16") = a
       
         .Range("A17:B17").Merge
      .Range("A17") = "Fecha Desembolso:"
       .Range("C17:E17").Merge
       .Range("C17") = gf_FormatoFecha(r_rst_Princi!HIPMAE_FECDES)
       
         .Range("A18:B18").Merge
      .Range("A18") = "Tasa de Interés:"
       .Range("C18:E18").Merge
       .Range("C18") = ((r_rst_Princi!HIPMAE_TASINT * 100) & " %")
       .Range("C18").NumberFormat = "###,###,##0.00"
        .Range("C18").HorizontalAlignment = xlHAlignLeft
       

      .Range("F14:G14").Merge
      .Range("F14") = "Código Cliente:"
       .Range("H14:I14").Merge
       .Range("H14") = "'" & r_rst_Princi!HIPMAE_NDOCLI
       
       
      .Range("H14").Font.Bold = True
      .Range("H14").Font.Underline = True
'       .Range("H15").NumberFormat = "@"
       
        .Range("F15:G15").Merge
       .Range("F15") = "Monto Préstamo:"
       .Range("H15:I15").Merge
       
'        "'" & Format(r_rst_Prindet!HIPCUO_CAPITA, "###,###,##0.00")
        .Range("H15") = "S/. " & Format(Trim(r_rst_Princi!HIPMAE_MTOPRE), "###,###,##0.00")
'        .Range("H15").HorizontalAlignment = xlHAlignCenter
'       .Range("H15") = "S/. " & Trim(r_rst_Princi!HIPMAE_MTOPRE)
'       .Range("H15").NumberFormat = "###,###,##0.00"
       
          .Range("F16:G16").Merge
          .Range("F16") = "Nro Cuotas:"
          .Range("H16:I16").Merge
          .Range("H16") = Trim(r_rst_Princi!HIPMAE_NUMCUO)
          
          .Range("H16").HorizontalAlignment = xlHAlignLeft
          
          
           .Range("F17:G17").Merge
          .Range("F17") = "Intereses Capitalizados:"
          .Range("H17:I17").Merge
          .Range("H17") = Trim(r_rst_Princi!HIPMAE_INTCAP)
          .Range("H17").NumberFormat = "###,###,##0.00"
           .Range("H17").HorizontalAlignment = xlHAlignLeft
          
          .Range("F18:G18").Merge
          .Range("F18") = "Monto Préstamo TNC:"
          .Range("H18:I18").Merge
          .Range("H18") = "S/. " & Trim(r_rst_Princi!HIPMAE_COSEFE)
       
 .Range(.Cells(20, 1), .Cells(20, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      '.CELDA(FILA, COLUMNA)
      
      .Cells(21, 1) = "Cuota"
      .Cells(21, 2) = "F.Vcto"
      .Cells(21, 3) = "Capital"
      .Cells(21, 4) = "Interés"
      .Cells(21, 5) = "S.Desg."
      .Cells(21, 6) = "S.Inm."
      .Cells(21, 7) = "Portes"
      .Cells(21, 8) = "T.Cuota"
      .Cells(21, 9) = "S.Capital"
      
       .Range(.Cells(21, 1), .Cells(21, 9)).HorizontalAlignment = xlHAlignCenter
      
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT CRE_HIPCUO.HIPCUO_NUMOPE, CRE_HIPCUO.HIPCUO_NUMCUO, HIPCUO_FECVCT, CRE_HIPCUO.HIPCUO_CAPITA, CRE_HIPCUO.HIPCUO_INTERE,"
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_DESORG, CRE_HIPCUO.HIPCUO_VIVORG, CRE_HIPCUO.HIPCUO_OTRORG, CRE_HIPCUO.HIPCUO_SALCAP, CRE_HIPMAE.HIPMAE_TDOCLI,"
      r_str_Parame = r_str_Parame & "CAST(CRE_HIPMAE.HIPMAE_NDOCLI AS VARCHAR(30)) AS HIPMAE_NDOCLI , CRE_HIPMAE.HIPMAE_PLAANO,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_CUOANO,CRE_HIPMAE.HIPMAE_PERGRA, CRE_HIPMAE.HIPMAE_NUMCUO, CRE_HIPMAE.HIPMAE_FECDES, CRE_HIPMAE.HIPMAE_MONEDA,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_MTOPRE, CRE_HIPMAE.HIPMAE_INTCAP, CRE_HIPMAE.HIPMAE_TASINT, CRE_HIPMAE.HIPMAE_COSEFE, CLI_DATGEN.DATGEN_APEPAT,"
      r_str_Parame = r_str_Parame & "CLI_DATGEN.DATGEN_APEMAT , CLI_DATGEN.DATGEN_APECAS, CLI_DATGEN.DATGEN_NOMBRE, CRE_PRODUC.PRODUC_DESCRI"
      r_str_Parame = r_str_Parame & " From "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO CRE_HIPCUO,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE CRE_HIPMAE,"
      r_str_Parame = r_str_Parame & "CLI_DATGEN CLI_DATGEN,"
      r_str_Parame = r_str_Parame & "CRE_PRODUC CRE_PRODUC"
      r_str_Parame = r_str_Parame & " Where "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_NUMOPE = CRE_HIPMAE.HIPMAE_NUMOPE AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_TDOCLI = CLI_DATGEN.DATGEN_TIPDOC AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_NDOCLI = CLI_DATGEN.DATGEN_NUMDOC AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_CODPRD = CRE_PRODUC.Produc_Codigo AND "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_TIPCRO = 1 AND "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_NUMOPE = '" & r_rst_Princi!HIPCUO_NUMOPE & "'"

'
'      MsgBox (r_str_Parame)
      
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Prindet, 3) Then
      Screen.MousePointer = 0
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Function
   End If
   
   If r_rst_Prindet.BOF And r_rst_Prindet.EOF Then
      r_rst_Prindet.Close
      Set r_rst_Prindet = Nothing
      Screen.MousePointer = 0
      Exit Function
   End If
      
      Dim r_int_NroFil, r_int_corre As Integer
      Dim sum1, sum2, sum3, sum4, sum5, sum6 As Double
       r_int_NroFil = 22
       r_int_corre = 1
      sum1 = 0
      sum2 = 0
      sum3 = 0
      sum4 = 0
      sum5 = 0
      sum6 = 0
      
    
      r_rst_Prindet.MoveFirst
      
     Do While Not r_rst_Prindet.EOF


             .Cells(r_int_NroFil, 1) = r_rst_Prindet!HIPCUO_NUMCUO
             .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
             .Cells(r_int_NroFil, 2) = "" & gf_FormatoFecha(r_rst_Prindet!HIPCUO_FECVCT)
'            .Cells(r_int_NroFil, 3) = r_rst_Prindet!HIPCUO_CAPITA
'            .Cells(r_int_NroFil, 3).NumberFormat = "###,###,##0.00"
             .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignLeft
             .Cells(r_int_NroFil, 3) = "'" & Format(r_rst_Prindet!HIPCUO_CAPITA, "###,###,##0.00")
             .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignLeft
'            .Cells(r_int_NroFil, 4) = r_rst_Prindet!HIPCUO_INTERE
'            .Cells(r_int_NroFil, 4).NumberFormat = "###,###,##0.00"
             .Cells(r_int_NroFil, 4) = "'" & Format(r_rst_Prindet!HIPCUO_INTERE, "###,###,##0.00")
             .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 5) = r_rst_Prindet!HIPCUO_DESORG
            .Cells(r_int_NroFil, 5).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 6) = r_rst_Prindet!HIPCUO_VIVORG
            .Cells(r_int_NroFil, 6).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 7) = r_rst_Prindet!HIPCUO_OTRORG
            .Cells(r_int_NroFil, 7).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 8) = r_rst_Prindet!HIPCUO_CAPITA
            .Cells(r_int_NroFil, 8).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 9) = r_rst_Prindet!HIPCUO_SALCAP
            .Cells(r_int_NroFil, 9).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignLeft
            .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 9)).HorizontalAlignment = xlHAlignCenter
            
        sum1 = sum1 + CDbl(.Cells(r_int_NroFil, 3))
         sum2 = sum2 + CDbl(.Cells(r_int_NroFil, 4))
          sum3 = sum3 + CDbl(.Cells(r_int_NroFil, 5))
          sum4 = sum4 + CDbl(.Cells(r_int_NroFil, 6))
         sum5 = sum5 + CDbl(.Cells(r_int_NroFil, 7))
          sum6 = sum6 + CDbl(.Cells(r_int_NroFil, 8))
                  
       r_int_corre = r_int_corre + 1
       r_int_NroFil = r_int_NroFil + 1
       

       r_rst_Prindet.MoveNext
       DoEvents
       

      Loop
      
      .Cells(r_int_NroFil, 2) = "TOTALES"
      .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignLeft
      
      
      .Cells(r_int_NroFil, 3) = sum1
      .Cells(r_int_NroFil, 3).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignLeft
      
      
      .Cells(r_int_NroFil, 4) = sum2
      .Cells(r_int_NroFil, 4).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignLeft
      
      
      .Cells(r_int_NroFil, 5) = sum3
      .Cells(r_int_NroFil, 5).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignLeft
      
      
      .Cells(r_int_NroFil, 6) = sum4
      .Cells(r_int_NroFil, 6).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignCenter
      
      
      .Cells(r_int_NroFil, 7) = sum5
      .Cells(r_int_NroFil, 7).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      
      
      .Cells(r_int_NroFil, 8) = sum6
      .Cells(r_int_NroFil, 8).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignLeft
      

   End With
   
      
   
  
   fs_GenExcNuevo = ""
'   fs_GenExcNuevo = "_49_9_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".PDF"

      fs_GenExcNuevo = r_rst_Princi!HIPCUO_NUMOPE & ".PDF"
   
   r_obj_Excel.ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, FileName:="C:/PDF4/" & fs_GenExcNuevo, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
   r_obj_Excel.ActiveWorkbook.Close SaveChanges:=False
   
   
   r_obj_Excel.Application.Quit
   Set r_obj_Excel = Nothing
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   Screen.MousePointer = 0
   
End Function


End Function

Private Sub cmd_Imprim_Click()
   Dim r_int_Cont As Integer
   
   If chk_VistaPrevia.Value = 1 Then
      If MsgBox("¿Está seguro de visualizar todos los reportes?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
         Exit Sub
      End If
   Else
      If MsgBox("¿Está seguro de imprimir todos los reportes?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
         Exit Sub
      End If
   End If

   For r_int_Cont = 1 To grd_Listado.Rows - 1
      If Val(grd_Listado.TextMatrix(r_int_Cont, 2)) > 0 Then
         Screen.MousePointer = 11
         
         Select Case grd_Listado.TextMatrix(r_int_Cont, 0)
            Case 1:
               Call fs_LiqDes(1)
               DoEvents
            Case 2:
               Call fs_HojRes
               DoEvents
            Case 3:
               Call fs_Cronog_MiCasita
               DoEvents
            Case 4:
               Call fs_Cronog_Mivivienda_NCoCli
               DoEvents
            Case 5:
               Call fs_Cronog_Mivivienda_ConCli
               DoEvents
            Case 7, 9:
               Call fs_Cronog_Mivivienda_ConMVi
               DoEvents
            Case 8:
               Call fs_Cronog_Mivivienda_NCoMVi
               DoEvents
            Case 10:
               Call fs_Cronog_Mivivienda_NCoCof
               DoEvents
            Case 11:
               Call fs_ComPag
               DoEvents
            Case 12
               If moddat_g_int_InsAct = 0 Then
                  MsgBox "No se puede emitir este formato, porque no se realizó ninguna operación.", vbInformation, modgen_g_str_NomPlt
                  Screen.MousePointer = 0
                  Exit Sub
               End If
               If moddat_g_int_InsAct = moddat_g_int_TipMon Then
                  MsgBox "No se puede emitir este formato, porque la Moneda de Compra-Venta es igual a la Moneda de Préstamo.", vbInformation, modgen_g_str_NomPlt
                  Screen.MousePointer = 0
                  Exit Sub
               End If
               Call fs_LiqTipoCambio
               DoEvents
            Case 13:
               Call fs_LiqDes(13)
               DoEvents
         End Select
         
         Screen.MousePointer = 0
      End If
   Next
   
   End Sub


' RAT 03202020  INICIO

Private Function fs_GenExcNuevo2(p_NumOpe As String) As String
        Dim r_rst_Princi         As ADODB.Recordset
        Dim r_rst_Prindet        As ADODB.Recordset
        Dim r_obj_Excel          As EXCEL.Application
        Dim r_int_NumFil         As Integer
        Dim r_str_Parame         As String
        Dim r_str_nom, r_str_ce  As String

        r_str_Parame = ""
        r_str_Parame = r_str_Parame & "SELECT RPT_FORDES.FORDES_DIRINM, RPT_FORDES.FORDES_DSTINM, RPT_FORDES.FORDES_MODALI, "
        r_str_Parame = r_str_Parame & "RPT_FORDES.FORDES_EMPSEG, RPT_FORDES.FORDES_TIPSEG, RPT_FORDES.FORDES_PORITF, "
        r_str_Parame = r_str_Parame & " RPT_FORDES.FORDES_GASTAS, RPT_FORDES.FORDES_GASNOT, RPT_FORDES.FORDES_BLQREG, "
        r_str_Parame = r_str_Parame & "CRE_HIPDES.HIPDES_NUMOPE, CRE_HIPMAE.HIPMAE_NUMSOL, CRE_HIPMAE.HIPMAE_TDOCLI, "
        r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_NDOCLI, CRE_HIPMAE.HIPMAE_PLAANO, CRE_HIPMAE.HIPMAE_CUOANO, "
        r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_PERGRA, CRE_HIPMAE.HIPMAE_NUMCUO, CRE_HIPMAE.HIPMAE_MONEDA, "
        r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_CVTDOL, CRE_HIPMAE.HIPMAE_APODOL, CRE_HIPMAE.HIPMAE_CVTSOL, "
        r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_APOSOL, CRE_HIPMAE.HIPMAE_MTOPRE, CRE_HIPMAE.HIPMAE_TASINT, "
        r_str_Parame = r_str_Parame & " CRE_HIPMAE.HIPMAE_FOIPRE, CRE_HIPMAE.HIPMAE_FOIVIV, CRE_HIPMAE.HIPMAE_COSEFE, "
        r_str_Parame = r_str_Parame & " CRE_HIPMAE.HIPMAE_TINTTN, CLI_DATGEN.DATGEN_APEPAT, CLI_DATGEN.DATGEN_APEMAT, "
        r_str_Parame = r_str_Parame & " CLI_DATGEN.DATGEN_APECAS, CLI_DATGEN.DATGEN_NOMBRE, CRE_PRODUC.PRODUC_DESCRI, "
        r_str_Parame = r_str_Parame & " TRA_EVALEG.EVALEG_MTOHIP, TRA_POLIZA.POLIZA_NUMDES, TRA_POLIZA.POLIZA_NUMVIV "
        r_str_Parame = r_str_Parame & " From "
        r_str_Parame = r_str_Parame & " RPT_FORDES RPT_FORDES, CRE_HIPDES CRE_HIPDES, CRE_HIPMAE CRE_HIPMAE, CLI_DATGEN CLI_DATGEN, CRE_PRODUC CRE_PRODUC, TRA_EVALEG TRA_EVALEG, TRA_POLIZA TRA_POLIZA "
        r_str_Parame = r_str_Parame & " Where "
        r_str_Parame = r_str_Parame & " RPT_FORDES.FORDES_NUMOPE = CRE_HIPDES.HIPDES_NUMOPE  AND "
        r_str_Parame = r_str_Parame & " RPT_FORDES.FORDES_NUMOPE = CRE_HIPMAE.HIPMAE_NUMOPE AND "
        r_str_Parame = r_str_Parame & " CRE_HIPMAE.HIPMAE_TDOCLI = CLI_DATGEN.DATGEN_TIPDOC AND "
        r_str_Parame = r_str_Parame & " CRE_HIPMAE.HIPMAE_NDOCLI = CLI_DATGEN.DATGEN_NUMDOC AND "
        r_str_Parame = r_str_Parame & " CRE_HIPMAE.HIPMAE_CODPRD = CRE_PRODUC.PRODUC_CODIGO AND "
        r_str_Parame = r_str_Parame & " CRE_HIPMAE.HIPMAE_NUMSOL = TRA_EVALEG.EVALEG_NUMSOL AND CRE_HIPMAE.HIPMAE_NUMSOL = TRA_POLIZA.POLIZA_NUMSOL"
        r_str_Parame = r_str_Parame & " AND CRE_HIPDES.HIPDES_NUMOPE = '" & p_NumOpe & "'"
   
         '   MsgBox (r_str_Parame)
   
         '   HIPCUO_TIPCRO
   
         'MsgBox (r_str_Parame)

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

         Set r_obj_Excel = New EXCEL.Application
             r_obj_Excel.SheetsInNewWorkbook = 1
             r_obj_Excel.Workbooks.Add
   
         With r_obj_Excel.ActiveSheet
   
        'MARGENES
         '.PageSetup.LeftMargin = Application.CentimetersToPoints(1.5)
         '.PageSetup.RightMargin = Application.CentimetersToPoints(0.4)
         '.PageSetup.TopMargin = Application.CentimetersToPoints(1)
         '.PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
         .Columns("A").ColumnWidth = 10
         .Columns("B").ColumnWidth = 10
         .Columns("C").ColumnWidth = 8
         .Columns("D").ColumnWidth = 8
         .Columns("E").ColumnWidth = 8
         .Columns("F").ColumnWidth = 10
         .Columns("G").ColumnWidth = 10
         .Columns("H").ColumnWidth = 8
         .Columns("I").ColumnWidth = 8
          .Range(.Cells(1, 1), .Cells(600, 12)).Font.Name = "Arial (Western)"
          .Range(.Cells(1, 1), .Cells(600, 12)).Font.Size = 11
          .Range(.Cells(1, 1), .Cells(600, 12)).RowHeight = 14
      
          .Pictures.Insert(g_str_RutLog & "\" & "image001.gif").Select
      
         .Cells(7, 1) = "HOJA RESUMEN"
         .Cells(7, 1).Font.Size = 14
         .Range(.Cells(7, 1), .Cells(7, 9)).Merge
         .Range(.Cells(7, 1), .Cells(7, 9)).Font.Bold = True
         .Range(.Cells(7, 1), .Cells(7, 9)).Font.Underline = True
         .Range(.Cells(7, 1), .Cells(7, 9)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(7, 1), .Cells(7, 9)).HorizontalAlignment = xlHAlignCenter

         .Rows(1).RowHeight = 1
         .Rows(8).RowHeight = 9
         .Rows(9).RowHeight = 5
         .Range(.Cells(9, 1), .Cells(9, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
         '.CELDA(FILA, COLUMNA)
         .Cells(2, 7) = "Nombre Reporte:"
         .Cells(2, 7).Font.Size = 8
         .Cells(2, 7).HorizontalAlignment = xlHAlignLeft
         .Cells(3, 7) = "Fecha Emisión:"
         .Cells(3, 7).Font.Size = 8
         .Cells(3, 7).HorizontalAlignment = xlHAlignLeft
         .Cells(4, 7) = "Hora Emisión:"
         .Cells(4, 7).Font.Size = 8
         .Cells(4, 7).HorizontalAlignment = xlHAlignLeft
         .Cells(5, 7) = "Página:"
         .Cells(5, 7).Font.Size = 8
         .Cells(5, 7).HorizontalAlignment = xlHAlignLeft
         .Cells(2, 8) = "RESUMEN"
         .Cells(2, 8).Font.Size = 8
         .Cells(2, 8).HorizontalAlignment = xlHAlignRight
         .Cells(3, 8) = Format(date, "dd/mm/yyyy")
         .Cells(3, 8).Font.Size = 8
         .Cells(3, 8).HorizontalAlignment = xlHAlignRight
         .Cells(4, 8) = Format(Time, "hh:mm:ss")
         .Cells(4, 8).Font.Size = 8
         .Cells(4, 8).HorizontalAlignment = xlHAlignRight
         .Cells(5, 8) = "1"
         .Cells(5, 8).Font.Size = 8
         .Cells(5, 8).HorizontalAlignment = xlHAlignRight

         .Range("A10:C10").Merge
         .Range("A10") = "Nro Operación:"
         .Range("A10").Font.Size = 12
         .Range("D10:F10").Merge
         .Range("D10") = Mid(r_rst_Princi!HIPDES_NUMOPE, 1, 3) & "-" & Mid(r_rst_Princi!HIPDES_NUMOPE, 4, 2) & "-" & Mid(r_rst_Princi!HIPDES_NUMOPE, 6, 5)
         .Range("D10").Font.Size = 10
      
         .Range("A11:C11").Merge
         .Range("A11") = "Nro Solicitud:"
         .Range("A11").Font.Size = 12
         .Range("D11:F11").Merge
         .Range("D11") = Mid(r_rst_Princi!hipmae_numsol, 1, 3) & "-" & Mid(r_rst_Princi!hipmae_numsol, 4, 3) & "-" & Mid(r_rst_Princi!hipmae_numsol, 7, 2) & "-" & Mid(r_rst_Princi!hipmae_numsol, 9, 4)
         .Range("D11").Font.Size = 10
      
         .Range("A12:C12").Merge
         .Range("A12") = "Doc Identidad:"
         .Range("A12").Font.Size = 12
         .Range("D12:G12").Merge
         .Range("D12") = IIf(r_rst_Princi!HIPMAE_TDOCLI = 1, "DNI " & Trim(r_rst_Princi!HIPMAE_NDOCLI), "CE " & Trim(r_rst_Princi!HIPMAE_NDOCLI))
         .Range("D12").Font.Size = 10
         
         If IsNull(Trim(r_rst_Princi!DATGEN_APECAS)) Then
             r_str_nom = Trim(r_rst_Princi!DATGEN_APEPAT) & " " & Trim(r_rst_Princi!DATGEN_APEMAT) & " " & Trim(r_rst_Princi!DATGEN_NOMBRE)
         Else
              If Len(Trim(r_rst_Princi!DATGEN_APECAS)) = 0 Then
                 r_str_nom = Trim(r_rst_Princi!DATGEN_APEPAT) & " " & Trim(r_rst_Princi!DATGEN_APEMAT) & " " & Trim(r_rst_Princi!DATGEN_NOMBRE)
              Else
                 r_str_nom = Trim(r_rst_Princi!DATGEN_APEPAT) & " " & Trim(r_rst_Princi!DATGEN_APEMAT) & " DE " & Trim(r_rst_Princi!DATGEN_APECAS) & " " & Trim(r_rst_Princi!DATGEN_NOMBRE)
              End If
         End If
      
        .Range("A13:C13").Merge
        .Range("A13") = "Nombre Cliente:"
        .Range("A13").Font.Size = 12
        .Range("D13:G13").Merge
        .Range("D13") = r_str_nom
        .Range("D13").Font.Size = 10
       
        .Range("A14:C14").Merge
        .Range("A14") = "Tipo de Crédito Hipotecario:"
        .Range("A14").Font.Size = 12
        .Range("D14:H14").Merge
        .Range("D14") = r_rst_Princi!PRODUC_DESCRI
        .Range("D14").Font.Size = 10
       
        .Range("A15:C15").Merge
        .Range("A15") = "Modalidad:"
        .Range("A15").Font.Size = 12
        .Range("D15:H15").Merge
        .Range("D15") = r_rst_Princi!FORDES_MODALI
        .Range("D15").Font.Size = 10
       

         .Range("A17:C17").Merge
         .Range("A17") = "Dirección Inmueble:"
         .Range("A17").Font.Size = 12
         .Range("D16:I18").Merge
    
          'r_obj_Excel.Visible = True
       
          .Range("D16") = Trim(r_rst_Princi!FORDES_DIRINM & "")
          .Range("D16").HorizontalAlignment = xlHAlignJustify
          .Range("D16").Font.Size = 9



          '.Range("D17:I17").Merge
          '.Range("D17") = r_rst_Princi!FORDES_DSTINM
          '.Range("D17").Font.Size = 10

           .Range("A20:C20").Merge
           .Range("A20") = "Valor Total de la Vivienda:"
           .Range("A20").Font.Size = 12
           .Range("D20:F20").Merge
           .Range("D20") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " " & IIf(r_rst_Princi!HIPMAE_MONEDA = 1, Format(r_rst_Princi!HIPMAE_CVTSOL, "###,###,##0.00"), Format(r_rst_Princi!HIPMAE_CVTDOL, "###,###,##0.00"))
           '.Range("D20").NumberFormat = "###,###,##0.00"
           .Range("D20").Font.Size = 10

           .Range("A21:C21").Merge
           .Range("A21") = "Cuota Inicial:"
           .Range("A21").Font.Size = 12
           .Range("D21:F21").Merge
           .Range("D21") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " " & IIf(r_rst_Princi!HIPMAE_MONEDA = 1, Format(r_rst_Princi!HIPMAE_APOSOL, "###,###,##0.00"), Format(r_rst_Princi!HIPMAE_APODOL, "###,###,##0.00"))
           .Range("D21").Font.Size = 10


           .Range("A22:C22").Merge
           .Range("A22") = "Monto del Crédito:"
           .Range("A22").Font.Size = 12
           .Range("D22:E22").Merge
           .Range("D22") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " " & Format(r_rst_Princi!HIPMAE_MTOPRE, "###,###,##0.00")
           .Range("D22").Font.Size = 10


           .Range("F22:G22").Merge
           .Range("F22") = "Moneda:"
           .Range("F22").Font.Size = 12
           .Range("H22:J22").Merge
           .Range("H22") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "NUEVOS SOLES", "DOLARES AMERICANOS")
           .Range("H22").Font.Size = 10
    
           .Range("A23:C23").Merge
           .Range("A23") = "Total Intereses Compensat:"
           .Range("A23").Font.Size = 12
           .Range("D23:F23").Merge
           .Range("D23") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " " & Format(r_rst_Princi!HIPMAE_TINTTN, "###,###,##0.00")
           .Range("D23").Font.Size = 10

           .Range("A24:C24").Merge
           .Range("A24") = "Plazo del Crédito:"
           .Range("A24").Font.Size = 12
           .Range("D24:F24").Merge
           .Range("D24") = r_rst_Princi!HIPMAE_PLAANO & " AÑOS"
           .Range("D24").Font.Size = 10
       
       
           .Range("A25:C25").Merge
           .Range("A25") = "Período Gracia:"
           .Range("A25").Font.Size = 12
           .Range("D25").Merge
           .Range("D25") = r_rst_Princi!HIPMAE_PERGRA & " Mes(es)"
           .Range("D25").Font.Size = 10
           .Range("E25").Font.Size = 6
           .Range("E25") = "'(1)"
       
           .Range("A26:C26").Merge
           .Range("A26") = "Número de Cuotas:"
           .Range("A26").Font.Size = 12
           .Range("D26").Merge
           .Range("D26") = r_rst_Princi!HIPMAE_NUMCUO
           .Range("D26").Font.Size = 10
           .Range("E26") = "'(2)"
           .Range("E26").Font.Size = 6
           
           
      
           .Range("F26:G26").Merge
           .Range("F26:G26") = "Periodicidad:"
           .Range("F26").Font.Size = 11
      
           .Range("H26:I26").Merge
           .Range("H26") = "Mensual"
           .Range("H26").Font.Size = 10
       
       
            If (r_rst_Princi!HIPMAE_CUOANO = 1) Then
                r_str_ce = "NO"
            ElseIf (r_rst_Princi!HIPMAE_CUOANO = 2) Then
                r_str_ce = "JULIO"
            ElseIf (r_rst_Princi!HIPMAE_CUOANO = 3) Then
                r_str_ce = "DICIEMBRE"
            ElseIf (r_rst_Princi!HIPMAE_CUOANO = 4) Then
                r_str_ce = "JULIO Y DICIEMBRE"
            End If
       
          .Range("A27:C27").Merge
          .Range("A27") = "Cuotas Extraordinarias:"
          .Range("A27").Font.Size = 12
          .Range("D27:E27").Merge
          .Range("D27") = r_str_ce
          .Range("D27").Font.Size = 10

          .Range("A28:C28").Merge
          .Range("A28") = "Garantía:"
          .Range("A28").Font.Size = 12
          .Range("D28:H28").Merge
          .Range("D28") = "Hipoteca hasta por la suma de: " & IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " " & Format(r_rst_Princi!EVALEG_MTOHIP, "###,###,##0.00")
          .Range("D28").Font.Size = 10
          .Range("I28") = "'(3)"
          .Range("I28").Font.Size = 6
       
          .Range("A29:B29").Merge
          .Range("A29") = "Tasas de Interés:"
          .Range("A29").Font.Size = 12
          .Range("C29") = "'(4)"
          .Range("C29").Font.Size = 6
              
          .Range("D29:I30").Merge
          .Range("D29") = "Tasa de Interés Compesatoria Efectiva Anual Fija de: " & Format(r_rst_Princi!HIPMAE_TASINT, "###,###,##0.00") & "%" & " " & "(Base 360)"
          .Range("D29").Font.Size = 10
           .Range("D29").HorizontalAlignment = xlHAlignJustify

          .Range("A31:C31").Merge
          .Range("A31") = "Tasa Costo Efectivo Anual:"
          .Range("A31").Font.Size = 12
          .Range("D31:I31").Merge
          .Range("D31") = Format(r_rst_Princi!HIPMAE_COSEFE, "###,###,##0.00") & "% " & "(Base 360)."
          .Range("D31").Font.Size = 10
           .Range("D31").HorizontalAlignment = xlHAlignJustify
       
          .Range("A32:B32").Merge
          .Range("A32") = "Impuestos :"
          .Range("A32").Font.Size = 12
          .Range("D32:F32").Merge
          .Range("D32") = "ITF: " & r_rst_Princi!FORDES_PORITF & "%"
          .Range("D32").Font.Size = 10
    
          .Range("A33:C33").Merge
          .Range("A33") = "Gastos:"
          .Range("A33").Font.Size = 12
          
          .Range("D34:G34").Merge
          .Range("D34") = "1. Gasto por Seguro de Desgravamen: "
          .Range("D34").Font.Size = 12
          .Range("H34") = "'(5)"
          .Range("H34").Font.Size = 6
          

          .Range("D35:F35").Merge
          .Range("D35:F35") = "Monto de la Prima:"
          .Range("D35").Font.Size = 11
          .Range("G35:I35").Merge
          .Range("G35") = r_rst_Princi!HIPMAE_FOIPRE & " %"
          .Range("G35").Font.Size = 9
           
          .Range("D36:F36").Merge
          .Range("D36") = "Compañía de Seguros:"
          .Range("D36").Font.Size = 11
          .Range("G36:I37").Merge
          .Range("G36").HorizontalAlignment = xlHAlignJustify
          .Range("G36") = r_rst_Princi!FORDES_EMPSEG
          .Range("G36").Font.Size = 9
        
          .Range("D38:F38").Merge
          .Range("D38") = "Tipo de Seguro:"
          .Range("D38").Font.Size = 11
          .Range("G38:I38").Merge
          .Range("G38") = r_rst_Princi!FORDES_TIPSEG
          .Range("G38").Font.Size = 9
          
          .Range("D39:F39").Merge
          .Range("D39") = "Nro. de Póliza:"
          .Range("D39").Font.Size = 11
          .Range("G39:I39").Merge
          .Range("G39") = r_rst_Princi!POLIZA_NUMDES
          .Range("G39").Font.Size = 9
         
                
         
          .Range("D40:F40").Merge
          .Range("D40") = "Riesgos de Cobertura:"
          .Range("D40").Font.Size = 11
          .Range("G40") = "'(6)"
          .Range("G40").Font.Size = 6
          .Range("H40:I40").Merge
          .Range("H40") = 0
          .Range("H40").Font.Size = 9
         
         
          
          .Range("D41:G41").Merge
          .Range("D41") = "2. Gasto por Seguro de Inmueble: "
          .Range("D41").Font.Size = 12
          .Range("H41") = "'(7)"
          .Range("H41").Font.Size = 6
          
          .Range("D42:F42").Merge
          .Range("D42") = "Monto de la Prima:"
          .Range("D42").Font.Size = 11
          .Range("G42:I42").Merge
          .Range("G42") = r_rst_Princi!HIPMAE_FOIVIV & " %"
          .Range("G42").Font.Size = 9
         
          .Range("D43:F43").Merge
          .Range("D43") = "Compañía de Seguros:"
          .Range("D43").Font.Size = 11
          .Range("G43:I44").Merge
          .Range("G43").HorizontalAlignment = xlHAlignJustify
          .Range("G43") = r_rst_Princi!FORDES_EMPSEG
          .Range("G43").Font.Size = 9
         
          .Range("D45:F45").Merge
          .Range("D45") = "Nro. de Póliza:"
          .Range("D45").Font.Size = 11
          .Range("G45:I45").Merge
          .Range("G45") = r_rst_Princi!POLIZA_NUMVIV
          .Range("G45").Font.Size = 9
         
         
          .Range("D46:F46").Merge
          .Range("D46") = "Riesgos de Cobertura:"
          .Range("D46").Font.Size = 11
          .Range("G46") = "'(8)"
           .Range("G46").Font.Size = 6
          .Range("H46:I46").Merge
          .Range("H46") = 0
          .Range("H46").Font.Size = 9
         
          .Range("D47:F47").Merge
          .Range("D47") = "3. Gasto por Tasación: "
          .Range("D47").Font.Size = 11
          .Range("G47:I47").Merge
          .Range("G47") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " " & Format(r_rst_Princi!FORDES_GASTAS, "###,###,##0.00")
          .Range("G47").Font.Size = 9
          
           .Range("D49:F49").Merge
           .Range("D49") = "4. Gasto Notariales: "
           .Range("D49").Font.Size = 11
           .Range("G49:H49").Merge
           .Range("G49") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " " & Format(r_rst_Princi!FORDES_GASNOT, "###,###,##0.00")
           .Range("G49").Font.Size = 9
           .Range("I49").Font.Size = 6
           .Range("I49") = "'(9)"

           .Range("D51:F51").Merge
           .Range("D51") = "5. Gasto Registrales: "
           .Range("D51").Font.Size = 11
           .Range("G51") = "'(10)"
           .Range("G51").Font.Size = 6

           .Range("D52:F52").Merge
           .Range("D52") = "Bloqueo Registral:"
           .Range("D52").Font.Size = 11
           .Range("G52:I52").Merge
           .Range("G52") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " " & Format(r_rst_Princi!FORDES_BLQREG, "###,###,##0.00")
           .Range("G52").Font.Size = 9


r_obj_Excel.Visible = True
           .Range("D53:F53").Merge
           .Range("D53") = "Compra Venta:"
           .Range("D53").Font.Size = 11
           .Range("G53:I54").Merge
           .Range("G53") = "Precio Venta*T.C.*0.03%+ S/33.00 por cada ficha"
           .Range("G53:I54").HorizontalAlignment = xlHAlignJustify
           
           .Range("G53").Font.Size = 9
           
           .Range("A54:C54").Merge
           .Range("A54") = "Penalidad:"
           .Range("A54").Font.Size = 11
           
           
           
           .Range("D55:H55").Merge
           .Range("D55") = "1. Pago Posterior a la fecha de vencimiento:"
           .Range("I55") = "'(11)"
           .Range("I55").Font.Size = 6
           

           .Range("D56:F56").Merge
           .Range("D56") = "A partir del dia 1:"
           .Range("D56").Font.Size = 11
           .Range("G56:I56").Merge
           .Range("G56") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " 30.00"
           .Range("G56").Font.Size = 11

            .Range("D57:F57").Merge
            .Range("D57") = "A partir del dia 8:"
            .Range("D57").Font.Size = 11
            .Range("G57:I57").Merge
            .Range("G57") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " 50.00"
            .Range("G57").Font.Size = 11
            
            .Range("D58:F58").Merge
            .Range("D58") = "A partir del dia 15:"
            .Range("D58").Font.Size = 11
            .Range("G58:I58").Merge
            .Range("G58") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "S/.", "US$") & " 100.00"
            .Range("G58").Font.Size = 11



            .Range("A60:I61").Merge
            .Range("A60") = "Ante el incumplimiento de pago de su cuota en la fecha prevista será reportado negativamente a las Centrales de Riesgo."
            .Range("A60:H61").HorizontalAlignment = xlJustify
            
      
             
            .Range("A63:I68").Merge
            .Range("A63") = "Declaro que la presente Hoja Resumen y el Contrato de Crédito Hipotecario que he suscrito con miCasita hipotecaria, así como el Cronograma de Pagos que forman parte integrante de esta Hoja Resumen, me han sido entregados para lectura, habiéndoseme (nos) absuelto todas las interrogantes formuladas, por lo que encontrándome (nos) conforme en todas las condiciones mencionadas, lo(s) firmo (amos) en dos ejemplares en señal de conformidad a las condiciones establecidas en los mismos."
             .Range("A63:I68").HorizontalAlignment = xlJustify
           
            .Range("A75:I75").Merge
            .Range("A75") = "San Isidro,____________ de _______________________________ del 20__________"
            .Range("A75:I75").HorizontalAlignment = xlJustify
           
             .Range("B86:E86").Merge
             .Range("B86") = "________________________"
             .Range("B87:E87").Merge
             .Range("B87") = "Firma del Cliente"
             
             .Range("G86:I86").Merge
             .Range("G86") = "________________________"
             .Range("G87:I87").Merge
             .Range("G87") = "Firma del Cónyuge"
             
             .Range("B99:E99").Merge
             .Range("B99") = "________________________"
             .Range("B100:E100").Merge
             .Range("B100") = "p/ miCasita Hipotecaria"
            
             .Range("G91:I91").Merge
             .Range("G91").Select
             .Pictures.Insert(g_str_RutLog & "\" & "image009.png").Select
             
             .Range("G99:I99").Merge
             .Range("G99") = "________________________"
             .Range("G100:I100").Merge
             .Range("G100") = "p/ miCasita Hipotecaria"
             
             
                          
              .Range("A105:I106").Merge
              .Range("A105") = "(1) Los intereses que se generen durante el período de gracia se capitalizarán y pasarán a formar parte del principal del crédito."
               .Range("A105:I106").HorizontalAlignment = xlJustify
               
            
               .Range("A107:I107").Merge
               .Range("A107") = "(2) Según detalle que consta en el Cronograma de Pagos que se entrega a El(s) Cliente(s)."
               .Range("A107").HorizontalAlignment = xlJustify
                
               .Range("A108:I108").Merge
               .Range("A108") = "(3) Según lo pactado en las Cláusulas Adicionales sobre Crédito Hipotecario.  "
             
               .Range("A109:I111").Merge
               .Range("A109") = "(4) Sobre la base de cálculo de 360 días calendarios. Las tarifas están sujetas a variación de acuerdo a lo establecido en el contrato correspondiente. Tanto las modificaciones como la determinación y recalculo de cuotas se efectuaran conforme a la legislación vigente."
               .Range("A109:I111").HorizontalAlignment = xlJustify
              
               .Range("A112:I112").Merge
               .Range("A112") = "(5) Según consta en la respectiva Póliza de Seguro entregada a El(s) Cliente(s)."
               
               
               .Range("A113:I113").Merge
               .Range("A113") = "(6) Riegos Cubiertos:"
               .Range("A114:I114").Merge
               .Range("A114") = "Muerte Natural: Cubre el fallecimiento del asegurado por causas naturales."
               
               .Range("A115:I115").Merge
               .Range("A115") = "Muerte Accidental: Cubre el fallecimiento del asegurado por causas accidentales."
               
               
               .Range("A116:I118").Merge
               .Range("A116") = "Invalidez Total Permanente Definitiva Por Enfermedad: Pérdida o disminución física o intelectual o superior a los 2/3 de su capacidad de trabajo, reconocida por la Compañía según el Dictamen de Evaluación y Calificación de la Invalidez total Permanente y Definitiva."
               .Range("A116:I118").HorizontalAlignment = xlJustify
              
               .Range("A119:I119").Merge
               .Range("A119") = "Invalidez Total Permanente Definitiva Por Accidente : Para los efectos de esta cobertura:"
               
               
               .Range("A120:I124").Merge
               .Range("A120") = "a) Fractura incurable de la columna vertebral; b) Descerebramiento que impida realizar trabajo alguno por el resto de su vida; c)Pérdida total o funcional absoluta de: La visión de ambos ojos, Ambos brazos o ambas manos, Ambas piernas o ambos pies, Una mano y un pie, siempre y cuando se determine una discapacidad superior o igual a los 2/3 de su capacidad de trabajo, reconocida por la Compañía según el Dictamen de Evaluación y Calificación de la Invalidez total Permanente."
               
              .Range("A120:I124").HorizontalAlignment = xlJustify
              
            
               .Range("A125:I125").Merge
               .Range("A125") = "(7) Según consta en la respectiva Póliza de Seguro entregada a El(s) Cliente(s)."
               .Range("A126:I126").Merge
               .Range("A126") = "(8) Riesgos cubiertos:"
               
               .Range("A127:I128").Merge
               .Range("A127") = "Todo riesgo de Incendio incluyendo terremoto, terrorismo y riesgos políticos: suma asegurada de hasta S/.400,000.00 equivalente en dólares del valor del crédito."
              .Range("A127:I128").HorizontalAlignment = xlJustify
               
                .Range("A129:I129").Merge
               .Range("A129") = "Terrorismo y riesgos políticos, al 100% de la Suma Asegurada: suma asegurada de hasta S/.400,000.00 equivalente en dólares del valor del crédito."
               .Range("A129:I129").HorizontalAlignment = xlJustify
               
               .Range("A131:I133").Merge
               .Range("A131") = "Gatos Extraordinarios, incluyendo honorarios profesionales, licencias, patentes de cualquier tipo, impuestos, defensa, salvamento, limpieza, remoción de escombros, documentos y modelos, costos de extinguir el incendio, reacondionamiento provisional: suma asegurada de S/.60,000.00 o US$ 20,000.00"
               .Range("A131:I133").HorizontalAlignment = xlJustify
        
               .Range("A134:I134").Merge
               .Range("A134") = "Gastos extras:  Suma asegurada de S/.30,000.00 o US$ 10,000.00"
               
               .Range("A135:I135").Merge
               .Range("A135") = "Rotura Accidental de Vidrios y/o cristales: suma asegurada de S/.15,000.00 o US$ 5,000.00"
               
               .Range("A136:I138").Merge
               .Range("A136") = "(9) Los gastos notariales adicionales no contemplados en el gasto de cierre serán de cuenta única y exclusiva de El(s) Cliente(s), miCasita hipotecaria podrá incorporar dichos gastos en el importe de las cuotas mensuales."
               .Range("A136:I138").HorizontalAlignment = xlJustify
            
               .Range("A139:I141").Merge
               .Range("A139") = "(10) Según tarifario de SUNARP. Los gastos registrales no contemplados en el gasto de cierre serán de cuenta única y exclusiva de El(s) Cliente(s), miCasita hipotecaria podrá incorporar dichos gastos en el importe de las cuotas mensuales."
              
                .Range("A139:I141").HorizontalAlignment = xlJustify
   
               .Range("A142:I142").Merge
               .Range("A142") = "(11) Corresponde a la penalidad por el incumplimiento de pago."
               
               .Range("A143:I146").Merge
               .Range("A143") = "(**) El(s) Cliente(s) tiene derecho a efectuar pagos anticipados o prepagos en forma total o parcial, con la consiguiente reducción de los intereses al día de pago. El(s) Cliente(s) debe comunicar el pago anticipado de sus cuotas expresamente y por escrito a miCasita hipotecaria."
              .Range("A143:I146").HorizontalAlignment = xlJustify
       
               .Range("A147:I148").Merge
               .Range("A147") = "(***) La realización de un prepago parcial o la cancelación total del crédito generará para El(s) Cliente(s) la pérdida de la parte proporcional que tenía el Tramo Concesional al momento del desembolso del Crédito."
               .Range("A147:I148").HorizontalAlignment = xlJustify
                 
               .Range("A150:I152").Merge
               .Range("A150") = "(****) Ante el incumplimiento del pago según las condiciones pactadas, se procederá a realizar el reporte correspondiente a la Central de Riesgos con la calificación que corresponda, de conformidad con el Reglamento para la Evaluación y Clasificación del Deudor."
               .Range("A150:I152").HorizontalAlignment = xlJustify
             
               .Range("A153:I157").Merge
               .Range("A153") = "(*****) Si producto de dolo o culpa debidamente acreditados, se induce a error a El(s) Cliente(s) y como consecuencia de ello éste realiza un pago en exceso, dicho monto es recuperable y devengará hasta su devolución el máximo de la suma por concepto de intereses compensatorios y moratorios que se hayan pactado para la operación crediticia o en su defecto, el interés legal."
               .Range("A153:I157").HorizontalAlignment = xlJustify
                 
            End With
   
                fs_GenExcNuevo2 = ""
                'fs_GenExcNuevo = "_49_9_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".PDF"

                fs_GenExcNuevo2 = r_rst_Princi!HIPDES_NUMOPE & Format(Time, "hhmmss") & ".PDF"
   
                r_obj_Excel.ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, FileName:="C:/PDFPRUEBAS3/" & fs_GenExcNuevo2, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
                r_obj_Excel.ActiveWorkbook.Close SaveChanges:=False

           r_obj_Excel.Application.Quit
           Set r_obj_Excel = Nothing
   
           r_rst_Princi.Close
           Set r_rst_Princi = Nothing
           Screen.MousePointer = 0
   
End Function

' RAT 03202020  FIN

