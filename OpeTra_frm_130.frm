VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Caj_GasCie_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   1845
   ClientTop       =   1860
   ClientWidth     =   14790
   Icon            =   "OpeTra_frm_130.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14835
      _Version        =   65536
      _ExtentX        =   26167
      _ExtentY        =   15425
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
         TabIndex        =   1
         Top             =   750
         Width           =   14745
         _Version        =   65536
         _ExtentX        =   26009
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
         Begin VB.CommandButton cmd_GasAdm 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_130.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Transferir Gastos de Cierre"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14100
            Picture         =   "OpeTra_frm_130.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salida"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   14745
         _Version        =   65536
         _ExtentX        =   26009
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
            Height          =   255
            Left            =   630
            TabIndex        =   5
            Top             =   60
            Width           =   7875
            _Version        =   65536
            _ExtentX        =   13891
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Operaciones Financieras"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   255
            Left            =   630
            TabIndex        =   6
            Top             =   330
            Width           =   7875
            _Version        =   65536
            _ExtentX        =   13891
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Transferencia de Gastos de Cierre"
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
            Left            =   14220
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
            Left            =   13260
            Top             =   90
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
            Left            =   12540
            Top             =   90
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_130.frx":0758
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7245
         Left            =   30
         TabIndex        =   7
         Top             =   1440
         Width           =   14745
         _Version        =   65536
         _ExtentX        =   26009
         _ExtentY        =   12779
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   1560
            TabIndex        =   9
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   3060
            TabIndex        =   10
            Top             =   60
            Width           =   4635
            _Version        =   65536
            _ExtentX        =   8176
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_Import 
            Height          =   285
            Left            =   8550
            TabIndex        =   11
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Imp. Asignado"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   6855
            Left            =   30
            TabIndex        =   12
            Top             =   360
            Width           =   14685
            _ExtentX        =   25903
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   30
            Cols            =   28
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_Moneda 
            Height          =   285
            Left            =   7680
            TabIndex        =   13
            Top             =   60
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   12840
            TabIndex        =   14
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Imp. Pagado"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   9960
            TabIndex        =   15
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud Ant."
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   11460
            TabIndex        =   16
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Pago"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frm_Caj_GasCie_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_PorITF     As Double

Private Sub cmd_GasAdm_Click()
   Dim r_str_CodPrd  As String
   Dim r_str_TIPMON  As String
   Dim r_str_NumSol  As String
   Dim r_str_TipDoc  As String
   Dim r_str_NumDoc  As String
   Dim r_str_CodBan  As String
   Dim r_str_NumCta  As String
   Dim r_str_NumCom  As String
   Dim r_str_SucMov  As String
   Dim r_str_FecMov  As String
   Dim r_str_NumMov  As String
   Dim r_dbl_ImpAsg  As Double
   Dim r_dbl_ImpTra  As Double
   Dim r_dbl_ImpPag  As Double
   Dim r_dbl_ImpITF  As Double
   Dim r_dbl_ImpTot  As Double
   Dim r_dbl_PorITF  As Double
   Dim r_str_Operac  As String
   Dim r_lng_NumMov  As Long
   Dim r_str_SolAnt  As String
   Dim r_lng_NMov01  As Long
   Dim r_lng_NMov02  As Long
   Dim r_str_NomCli  As String
   Dim r_int_CodIns  As Integer
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 2
   r_str_NomCli = grd_Listad.Text
   
   grd_Listad.Col = 4
   r_dbl_ImpAsg = CDbl(grd_Listad.Text)
   
   grd_Listad.Col = 7
   r_dbl_ImpTra = CDbl(grd_Listad.Text)
   
   grd_Listad.Col = 8
   r_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 10
   r_str_NumSol = grd_Listad.Text
   
   grd_Listad.Col = 11
   r_str_TIPMON = grd_Listad.Text
   
   grd_Listad.Col = 12
   moddat_g_str_CodConHip = grd_Listad.Text
   
   grd_Listad.Col = 13
   moddat_g_str_CodEjeSeg = grd_Listad.Text
   
   grd_Listad.Col = 14
   r_int_CodIns = CInt(grd_Listad.Text)
   
   grd_Listad.Col = 15
   r_str_TipDoc = grd_Listad.Text
   
   grd_Listad.Col = 16
   r_str_NumDoc = grd_Listad.Text
   
   grd_Listad.Col = 17
   r_str_CodBan = grd_Listad.Text
   
   grd_Listad.Col = 18
   r_str_NumCta = grd_Listad.Text
   
   grd_Listad.Col = 19
   r_str_NumCom = grd_Listad.Text
   
   grd_Listad.Col = 20
   r_dbl_ImpPag = CDbl(grd_Listad.Text)
   
   grd_Listad.Col = 21
   r_dbl_ImpITF = CDbl(grd_Listad.Text)
   
   grd_Listad.Col = 22
   r_dbl_ImpTot = CDbl(grd_Listad.Text)
   
   grd_Listad.Col = 23
   r_dbl_PorITF = CDbl(grd_Listad.Text)
   
   grd_Listad.Col = 23
   r_dbl_PorITF = CDbl(grd_Listad.Text)
   
   grd_Listad.Col = 24
   r_str_SucMov = grd_Listad.Text
   
   grd_Listad.Col = 25
   r_str_FecMov = grd_Listad.Text
   
   grd_Listad.Col = 26
   r_str_NumMov = grd_Listad.Text
   
   grd_Listad.Col = 27
   r_str_SolAnt = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If r_dbl_ImpAsg <> r_dbl_ImpTra Then
      MsgBox "El Importe Asignado y el Importe a Trasladar no son iguales. ", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If MsgBox("¿Está seguro de realizar el Traslado de Gastos de Cierre?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
'**INICIO COMENTA EL ASIENTO DE EXTORNO DE GASTO DE CIERRE
'   'Registrando Extorno de Movimiento Anterior
'   r_str_Operac = moddat_gf_Consulta_Operac(r_str_CodPrd, "210")
'   r_str_Operac = r_str_TIPMON & Right(r_str_Operac, 5)
'
'   'Obteniendo Número de Movimiento
'   r_lng_NumMov = opecaj_gf_Genera_NumMov()
'   r_lng_NMov01 = r_lng_NumMov
'
'   'Registrando Movimiento
'   If Not opecaj_gf_Inserta_CajMov(modgen_g_str_CodUsu, "2101", r_str_SolAnt, "", CInt(r_str_TipDoc), r_str_NumDoc, r_str_CodBan, Format(date, "yyyymmdd"), _
'                                   r_str_NumCta, "0", CInt(r_str_TIPMON), r_dbl_ImpPag, 1, modgen_g_str_CodSuc, 0, 0, 0, r_dbl_PorITF, r_dbl_ImpITF, _
'                                   r_dbl_IMPTOT, 1, r_str_NumMov, r_str_Operac, r_lng_NumMov, 1, "0", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, r_str_SucMov, r_str_FecMov) Then
'      Exit Sub
'   End If
'
'   'Registrando Movimiento Nuevo
'   r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "210")
'   r_str_Operac = CStr(moddat_g_int_TipMon) & Right(r_str_Operac, 5)
'
'   'Obteniendo Número de Movimiento
'   r_lng_NumMov = opecaj_gf_Genera_NumMov()
'   r_lng_NMov02 = r_lng_NumMov
'
'   'Registrando Movimiento
'   If Not opecaj_gf_Inserta_CajMov(modgen_g_str_CodUsu, "1101", r_str_NumSol, "", CInt(r_str_TipDoc), r_str_NumDoc, r_str_CodBan, Format(date, "yyyymmdd"), _
'                                   r_str_NumCta, r_str_NumCom, CInt(r_str_TIPMON), r_dbl_ImpPag, 0, modgen_g_str_CodSuc, 0, 0, 0, r_dbl_PorITF, r_dbl_ImpITF, _
'                                   r_dbl_IMPTOT, 0, "0", r_str_Operac, r_lng_NumMov, 1, "0", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0) Then
'      Exit Sub
'   End If
'**FIN COMENTA EL ASIENTO DE EXTORNO DE GASTO DE CIERRE

   'Buscando Gastos de Cierre
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_GASADM "
   g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & r_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "   AND GASADM_SITUAC = 2"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      If Not opecaj_gf_Pago_GasAdm(r_str_NumSol, g_rst_Genera!GASADM_CODGAS, CInt(r_str_TIPMON), g_rst_Genera!GASADM_IMPORT, r_dbl_PorITF, Format(date, "yyyymmdd"), r_str_Operac) Then
         Exit Sub
      End If
      g_rst_Genera.MoveNext
   Loop
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
      
  '----INICIO actualizar pago a proveedores
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.GASADM_FECPAGPRV,A.GASADM_TIPPAGPRV, A.GASADM_NUMDOCPRV, A.GASADM_MTOPAGPRV,  "
   g_str_Parame = g_str_Parame & "        A.GASADM_FECENTPRV, A.GASADM_NROCNT, A.GASADM_CODOPE, A.GASADM_CODGAS  "
   g_str_Parame = g_str_Parame & "   FROM TRA_GASADM A  "
   g_str_Parame = g_str_Parame & "  WHERE A.GASADM_NUMSOL = '" & Trim(r_str_SolAnt) & "'  "
   'g_str_Parame = g_str_Parame & "    AND A.GASADM_MTOPAGPRV > 0  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
    
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " UPDATE TRA_GASADM SET  "
      If Trim(g_rst_Genera!GASADM_FECPAGPRV & "") = "" Then
         g_str_Parame = g_str_Parame & "        GASADM_FECPAGPRV = null,  "
      Else
         g_str_Parame = g_str_Parame & "        GASADM_FECPAGPRV = " & g_rst_Genera!GASADM_FECPAGPRV & ",  "
      End If
      If Trim(g_rst_Genera!GASADM_TIPPAGPRV & "") = "" Then
         g_str_Parame = g_str_Parame & "        GASADM_TIPPAGPRV = null,  "
      Else
         g_str_Parame = g_str_Parame & "        GASADM_TIPPAGPRV = " & g_rst_Genera!GASADM_TIPPAGPRV & ",  "
      End If
      g_str_Parame = g_str_Parame & "        GASADM_NUMDOCPRV = '" & g_rst_Genera!GASADM_NUMDOCPRV & "',  "
      If Trim(g_rst_Genera!GASADM_MTOPAGPRV & "") = "" Then
         g_str_Parame = g_str_Parame & "        GASADM_MTOPAGPRV = null,  "
      Else
         g_str_Parame = g_str_Parame & "        GASADM_MTOPAGPRV = " & g_rst_Genera!GASADM_MTOPAGPRV & ",  "
      End If
      If Trim(g_rst_Genera!GASADM_FECENTPRV & "") = "" Then
         g_str_Parame = g_str_Parame & "        GASADM_FECENTPRV = null,  "
      Else
         g_str_Parame = g_str_Parame & "        GASADM_FECENTPRV = " & g_rst_Genera!GASADM_FECENTPRV & ",  "
      End If
      g_str_Parame = g_str_Parame & "        GASADM_NROCNT = '" & g_rst_Genera!GASADM_NROCNT & "'  "
      'g_str_Parame = g_str_Parame & "        GASADM_CODOPE = " & g_rst_Genera!GASADM_CODOPE & "  "
      g_str_Parame = g_str_Parame & "  WHERE GASADM_NUMSOL = '" & r_str_NumSol & "' "
      g_str_Parame = g_str_Parame & "    AND GASADM_CODGAS = " & g_rst_Genera!GASADM_CODGAS
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
         Exit Sub
      End If
         
      g_rst_Genera.MoveNext
   Loop
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   '----FIN actualizar pago a proveedores
   
   
   'Actualizando en Seguimiento de Tasacion Pago de Gastos Administrativos
   If Not moddat_gf_Inserta_SegDet(r_str_NumSol, r_int_CodIns, 25, 0, "", 0, 0) Then
      Exit Sub
   End If
      
   'Enviando Correo Electrónico
   modgen_g_str_Mail_Asunto = "PAGO DE GASTOS DE CIERRE (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & gf_Formato_NumSol(r_str_NumSol) & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & r_str_TipDoc & "-" & r_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & r_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   
   Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False)
   
   'Impresion de Comprobantes
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
   Call opecaj_gs_ComPago(modgen_g_str_CodSuc, CStr(r_lng_NMov01), Format(date, "yyyymmdd"), 1, 1)
   Call opecaj_gs_ComPago(modgen_g_str_CodSuc, CStr(r_lng_NMov02), Format(date, "yyyymmdd"), 1, 2)
   Screen.MousePointer = 0
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "RPT_COMPGC"
   crp_Imprim.DataFiles(1) = "RPT_COMPGD"
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_COMPAG_01.RPT"
   crp_Imprim.SelectionFormula = "{RPT_COMPGC.COMPGC_CODTER} = '" & modgen_g_str_NombPC & "'"
   
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
      
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1495
   grd_Listad.ColWidth(1) = 1505
   grd_Listad.ColWidth(2) = 4625
   grd_Listad.ColWidth(3) = 875
   grd_Listad.ColWidth(4) = 1405
   grd_Listad.ColWidth(5) = 1495
   grd_Listad.ColWidth(6) = 1395
   grd_Listad.ColWidth(7) = 1495
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 0
   grd_Listad.ColWidth(14) = 0
   grd_Listad.ColWidth(15) = 0
   grd_Listad.ColWidth(16) = 0
   grd_Listad.ColWidth(17) = 0
   grd_Listad.ColWidth(18) = 0
   grd_Listad.ColWidth(19) = 0
   grd_Listad.ColWidth(20) = 0
   grd_Listad.ColWidth(21) = 0
   grd_Listad.ColWidth(22) = 0
   grd_Listad.ColWidth(23) = 0
   grd_Listad.ColWidth(24) = 0
   grd_Listad.ColWidth(25) = 0
   grd_Listad.ColWidth(26) = 0
   grd_Listad.ColWidth(27) = 0
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter

   'Obteniendo ITF
   l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
End Sub

Private Sub fs_Buscar()
Dim r_dbl_ITFGas     As Double
Dim r_str_SucMov     As String
Dim r_str_FecMov     As String
Dim r_str_NumMov     As String
Dim r_str_NueSol     As String
Dim r_int_MonPag     As Integer
Dim r_dbl_ImpPag     As Double
Dim r_dbl_ImpITF     As Double
Dim r_dbl_ImpTot     As Double
Dim r_dbl_NueGas     As Double
Dim r_int_TipDoc     As Integer
Dim r_str_NumDoc     As String
Dim r_str_FecPag     As String
Dim r_rst_Genera     As ADODB.Recordset
Dim r_rst_CajMov     As ADODB.Recordset
Dim r_str_CodBan     As String
Dim r_str_NumCta     As String
Dim r_str_NumCom     As String
Dim r_dbl_PorITF     As Double
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_CODPRD, SOLMAE_CODSUB, SOLMAE_TIPMON, SOLMAE_CONHIP, SOLMAE_EJESEG, SOLMAE_CODINS "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND (SOLMAE_CODINS = 31 OR SOLMAE_CODINS = 32) "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT SUM(GASADM_IMPORT) AS TOTGAS "
         g_str_Parame = g_str_Parame & "  FROM TRA_GASADM "
         g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' "
         g_str_Parame = g_str_Parame & "   AND GASADM_SITUAC = 2"
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
             Exit Sub
         End If
         
         'Si cliente no tiene Gastos Pagados
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) And Not IsNull(g_rst_Genera!TOTGAS) Then
            r_str_SucMov = ""
            r_str_FecMov = ""
            r_str_NumMov = ""
            r_int_MonPag = 0
            r_dbl_ImpPag = 0
            r_dbl_ImpITF = 0
            r_dbl_ImpTot = 0
            r_str_NueSol = ""
            r_dbl_NueGas = 0
            r_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
            r_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO & "")
            r_str_NueSol = g_rst_Princi!SOLMAE_NUMERO
            
            'Buscando Ultima Solicitud Rechazada
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "SELECT * FROM CRE_SOLMAE "
            g_str_Parame = g_str_Parame & " WHERE SOLMAE_SITUAC = 3 "
            g_str_Parame = g_str_Parame & "   AND SOLMAE_TITTDO = " & CStr(r_int_TipDoc) & " "
            g_str_Parame = g_str_Parame & "   AND SOLMAE_TITNDO = '" & r_str_NumDoc & "' "
            g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO DESC "
            
            If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
               Exit Sub
            End If
         
            If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
               r_rst_Genera.MoveFirst
               'r_str_AntSol = Trim(r_rst_Genera!SOLMAE_NUMERO)
               
               'Buscar Pago de Gastos de Cierre que no han sido Reversados
               g_str_Parame = ""
               g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV "
               g_str_Parame = g_str_Parame & " WHERE CAJMOV_TIPMOV = 1101 "
               g_str_Parame = g_str_Parame & "   AND CAJMOV_NUMOPE = '" & r_rst_Genera!SOLMAE_NUMERO & "' "
               g_str_Parame = g_str_Parame & "   AND CAJMOV_FLGREV = 0 "
               
               If Not gf_EjecutaSQL(g_str_Parame, r_rst_CajMov, 3) Then
                   Exit Sub
               End If
               
               If Not (r_rst_CajMov.BOF And r_rst_CajMov.EOF) Then
                  r_str_SucMov = r_rst_CajMov!CAJMOV_SUCMOV
                  r_str_FecMov = CStr(r_rst_CajMov!CAJMOV_FECMOV)
                  r_str_NumMov = CStr(r_rst_CajMov!CAJMOV_NUMMOV)
                  r_str_FecPag = CStr(r_rst_CajMov!CAJMOV_FECDEP)
                  r_str_CodBan = Trim(r_rst_CajMov!CAJMOV_CODBAN & "")
                  r_str_NumCta = Trim(r_rst_CajMov!CAJMOV_NUMCTA & "")
                  r_str_NumCom = Trim(r_rst_CajMov!CAJMOV_NUMCOM & "")
                  r_int_MonPag = r_rst_CajMov!CAJMOV_MONPAG
                  r_dbl_ImpPag = r_rst_CajMov!CAJMOV_IMPPAG
                  r_dbl_ImpITF = r_rst_CajMov!CAJMOV_ITFIMP
                  r_dbl_ImpTot = r_rst_CajMov!CAJMOV_IMPTOT
                  r_dbl_PorITF = r_rst_CajMov!CAJMOV_ITFPOR
               
                  'Asignando Grid
                  grd_Listad.Rows = grd_Listad.Rows + 1
                  grd_Listad.Row = grd_Listad.Rows - 1
                  
                  grd_Listad.Col = 0
                  grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
                  
                  grd_Listad.Col = 1
                  grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
                  
                  grd_Listad.Col = 2
                  grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
                  
                  grd_Listad.Col = 3
                  grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(r_int_MonPag))
                  
                  r_dbl_ITFGas = CDbl(gf_NueImp_Numero(gf_Truncar_Numero(g_rst_Genera!TOTGAS * (l_dbl_PorITF / 100), 2)))
                                    
                  grd_Listad.Col = 4
                  grd_Listad.Text = Format(g_rst_Genera!TOTGAS + r_dbl_ITFGas, "###,###,##0.00")
                  
                  grd_Listad.Col = 5
                  grd_Listad.Text = gf_Formato_NumSol(Trim(r_rst_Genera!SOLMAE_NUMERO))
                  
                  grd_Listad.Col = 6
                  grd_Listad.Text = gf_FormatoFecha(r_str_FecPag)
                  
                  grd_Listad.Col = 7
                  grd_Listad.Text = Format(r_dbl_ImpTot, "###,##0.00")
                  
                  grd_Listad.Col = 8
                  grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
                  
                  grd_Listad.Col = 9
                  grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
                  
                  grd_Listad.Col = 10
                  grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_NUMERO & "")
                  
                  grd_Listad.Col = 11
                  grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TIPMON)
               
                  grd_Listad.Col = 12
                  grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_CONHIP)
               
                  grd_Listad.Col = 13
                  grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_EJESEG)
               
                  grd_Listad.Col = 14
                  grd_Listad.Text = g_rst_Princi!SOLMAE_CODINS
               
                  grd_Listad.Col = 15
                  grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO)
                  
                  grd_Listad.Col = 16
                  grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_TITNDO)
                  
                  grd_Listad.Col = 17
                  grd_Listad.Text = r_str_CodBan
                  
                  grd_Listad.Col = 18
                  grd_Listad.Text = r_str_NumCta
                  
                  grd_Listad.Col = 19
                  grd_Listad.Text = r_str_NumCom
                  
                  grd_Listad.Col = 20
                  grd_Listad.Text = CStr(r_dbl_ImpPag)
                  
                  grd_Listad.Col = 21
                  grd_Listad.Text = CStr(r_dbl_ImpITF)
                  
                  grd_Listad.Col = 22
                  grd_Listad.Text = CStr(r_dbl_ImpTot)
                  
                  grd_Listad.Col = 23
                  grd_Listad.Text = CStr(r_dbl_PorITF)
                  
                  grd_Listad.Col = 24
                  grd_Listad.Text = r_str_SucMov
                  
                  grd_Listad.Col = 25
                  grd_Listad.Text = r_str_FecMov
                  
                  grd_Listad.Col = 26
                  grd_Listad.Text = r_str_NumMov
                  
                  grd_Listad.Col = 27
                  grd_Listad.Text = r_rst_Genera!SOLMAE_NUMERO
               End If
            
               r_rst_CajMov.Close
               Set r_rst_CajMov = Nothing
            End If
            
            r_rst_Genera.Close
            Set r_rst_Genera = Nothing
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   If grd_Listad.Rows = 0 Then
      cmd_GasAdm.Enabled = False
      MsgBox "No se encontraron Solicitudes para Traslado de Gastos de Cierre.", vbInformation, modgen_g_str_NomPlt
   Else
      'Ordenando por Nombres de Clientes
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_GasAdm_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Len(Trim(pnl_Tit_NumSol.Tag)) = 0 Or pnl_Tit_NumSol.Tag = "D" Then
      pnl_Tit_NumSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_Moneda_Click()
   If Len(Trim(pnl_Tit_Moneda.Tag)) = 0 Or pnl_Tit_Moneda.Tag = "D" Then
      pnl_Tit_Moneda.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_Moneda.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_Import_Click()
   If Len(Trim(pnl_Tit_Import.Tag)) = 0 Or pnl_Tit_Import.Tag = "D" Then
      pnl_Tit_Import.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "N")
   Else
      pnl_Tit_Import.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "N-")
   End If
End Sub

