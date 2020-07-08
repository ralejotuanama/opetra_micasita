VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_Cofide_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15105
   Icon            =   "OpeTra_frm_818.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15090
      _Version        =   65536
      _ExtentX        =   26617
      _ExtentY        =   12091
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
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_818.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   1800
            Picture         =   "OpeTra_frm_818.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14370
            Picture         =   "OpeTra_frm_818.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_818.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_818.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Cliente"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   10680
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
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
            Width           =   8835
            _Version        =   65536
            _ExtentX        =   15584
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Generación Cartas COFIDE"
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
            Left            =   120
            Picture         =   "OpeTra_frm_818.frx":11AE
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   5415
         Left            =   30
         TabIndex        =   7
         Top             =   1410
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
         _ExtentY        =   9551
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   4995
            Left            =   60
            TabIndex        =   8
            Top             =   360
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   8811
            _Version        =   393216
            Rows            =   45
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   4680
            TabIndex        =   9
            Top             =   60
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DOI Cliente"
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
            Left            =   735
            TabIndex        =   10
            Top             =   60
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7064
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cliente"
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   10410
            TabIndex        =   11
            Top             =   60
            Width           =   2790
            _Version        =   65536
            _ExtentX        =   4921
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
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
         Begin Threed.SSPanel pnl_Tit_ParReg 
            Height          =   285
            Left            =   8520
            TabIndex        =   12
            Top             =   60
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Partida Registral"
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
         Begin Threed.SSPanel pnl_Tit_CodFMV 
            Height          =   285
            Left            =   6360
            TabIndex        =   14
            Top             =   60
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código FMV"
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
         Begin Threed.SSPanel SSPanel81 
            Height          =   285
            Left            =   13200
            TabIndex        =   16
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "  Seleccionar"
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
            Alignment       =   1
            Begin VB.CheckBox chkSeleccionar 
               BackColor       =   &H00004000&
               Caption         =   "Check1"
               Height          =   255
               Left            =   1110
               TabIndex        =   17
               Top             =   0
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_Tit_Item 
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Item"
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
Attribute VB_Name = "frm_Rpt_Cofide_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 6) = ""
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 6) = "X"
         Next r_Fila
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub cmd_Borrar_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer

   'valida selección
   r_int_ConSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 6) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionado Solicitudes para Eliminar.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirma
   If MsgBox("¿Está seguro de Eliminar las solicitudes seleccionadas?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If grd_Listad.Rows = 1 Then
      Call gs_LimpiaGrid(grd_Listad)
      fs_Activa (False)
   Else
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         If r_int_Contad = grd_Listad.Rows Then
            Exit For
         End If
         If r_int_Contad = grd_Listad.Row And r_int_Contad = 0 Then
               Call gs_LimpiaGrid(grd_Listad)
               fs_Activa (False)
               Exit For
         End If
         If grd_Listad.TextMatrix(r_int_Contad, 6) = "X" Then
            grd_Listad.RemoveItem (r_int_Contad)
            r_int_Contad = r_int_Contad - 1
         End If
      Next
    End If
    chkSeleccionar.Value = False
End Sub

Private Sub cmd_BusCli_Click()
   moddat_g_int_FlgCre = 3
   frm_Ges_CreHip_01.Show 1
End Sub

Private Sub cmd_Imprim_Click()
Dim r_str_Parame     As String
Dim r_int_NroFil     As Integer
Dim r_str_NumCar     As String

   If grd_Listad.Rows = 0 Then
      MsgBox "No se encontraron Solicitudes para imprimir.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
     
   'Confirma
   If MsgBox("¿Está seguro de Imprimir las solicitudes ?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
       
   'Elimina datos de la tabla temporal
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " DELETE FROM RPT_TABLA_TEMP "
   g_str_Parame = g_str_Parame & "  WHERE RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "    AND RPT_NOMBRE = 'REPORTE CONSTITUCION HIPOTECAS' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   'Grabacion en Tabla de Temporal
   For r_int_NroFil = 0 To grd_Listad.Rows - 1
   
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO RPT_TABLA_TEMP("
      g_str_Parame = g_str_Parame & "  RPT_PERMES, "
      g_str_Parame = g_str_Parame & "  RPT_PERANO, "
      g_str_Parame = g_str_Parame & "  RPT_TERCRE, "
      g_str_Parame = g_str_Parame & "  RPT_USUCRE, "
      g_str_Parame = g_str_Parame & "  RPT_NOMBRE, "
      g_str_Parame = g_str_Parame & "  RPT_MONEDA, "
      g_str_Parame = g_str_Parame & "  RPT_FECCRE, "
      g_str_Parame = g_str_Parame & "  RPT_HORCRE, "
      g_str_Parame = g_str_Parame & "  RPT_CODIGO, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD01, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD02, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD03, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD04, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD05, "
      g_str_Parame = g_str_Parame & "  RPT_VALCAD06, "
      g_str_Parame = g_str_Parame & "  RPT_VALNUM01) "
      
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "" & Month(date) & ", "
      g_str_Parame = g_str_Parame & "" & Year(date) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'REPORTE CONSTITUCION HIPOTECAS', "
      g_str_Parame = g_str_Parame & "'1', "
      g_str_Parame = g_str_Parame & "'" & Format(date, "DDMMYYYY") & "', "
      g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
      g_str_Parame = g_str_Parame & "" & r_int_NroFil & ", "
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_NroFil, 1) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_NroFil, 2) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_NroFil, 3) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_NroFil, 4) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_NroFil, 5) & "', "
      g_str_Parame = g_str_Parame & "'" & grd_Listad.TextMatrix(r_int_NroFil, 7) & "', "
      g_str_Parame = g_str_Parame & "" & r_int_NroFil + 1 & ")"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         Exit Sub
      End If
      
      'busca si tiene conyuge
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT A.HIPMAE_NUMOPE AS OPERACION, "
      r_str_Parame = r_str_Parame & "        TRIM(TRIM(G.DATGEN_APEPAT) || ' ' || TRIM(G.DATGEN_APEMAT) || ' ' ||  TRIM(G.DATGEN_NOMBRE)) AS CONYUGE, "
      r_str_Parame = r_str_Parame & "        CASE WHEN INSTR(TRIM(H.PARDES_DESCRI),'EXTRANJERIA') > 0 THEN 'CE' ELSE TRIM(H.PARDES_DESCRI) END || '-' || TRIM(G.DATGEN_NUMDOC) AS DOCIDE_CONYUGE "
      r_str_Parame = r_str_Parame & "   FROM CRE_HIPMAE A "
      r_str_Parame = r_str_Parame & "         LEFT JOIN CLI_DATGEN G ON G.DATGEN_TIPDOC = A.HIPMAE_TDOCYG AND G.DATGEN_NUMDOC = A.HIPMAE_NDOCYG "
      r_str_Parame = r_str_Parame & "         LEFT JOIN MNT_PARDES H ON H.PARDES_CODGRP = 230 AND H.PARDES_CODITE = G.DATGEN_TIPDOC "
      r_str_Parame = r_str_Parame & "  WHERE A.HIPMAE_NUMOPE = '" & grd_Listad.TextMatrix(r_int_NroFil, 7) & "' "
      
      If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
      Else
         If Not IsNull(g_rst_Genera!CONYUGE) Then
             g_str_Parame = ""
             g_str_Parame = g_str_Parame & "UPDATE RPT_TABLA_TEMP SET "
             g_str_Parame = g_str_Parame & "  RPT_VALCAD01 = RPT_VALCAD01 || chr(13) || '" & g_rst_Genera!CONYUGE & "' ,"
             g_str_Parame = g_str_Parame & "  RPT_VALCAD02 = RPT_VALCAD02 || chr(13) || '" & g_rst_Genera!DOCIDE_CONYUGE & "'"
             g_str_Parame = g_str_Parame & " WHERE RPT_PERMES = " & Month(date) & ""
             g_str_Parame = g_str_Parame & "   AND RPT_PERANO = " & Year(date) & ""
             g_str_Parame = g_str_Parame & "   AND RPT_TERCRE = '" & modgen_g_str_NombPC & "'"
             g_str_Parame = g_str_Parame & "   AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "'"
             g_str_Parame = g_str_Parame & "   AND RPT_NOMBRE = 'REPORTE CONSTITUCION HIPOTECAS'"
             g_str_Parame = g_str_Parame & "   AND RPT_MONEDA = 1 "
             g_str_Parame = g_str_Parame & "   AND TRIM(RPT_VALCAD06) =  '" & g_rst_Genera!OPERACION & "'"
         
             If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                Exit Sub
             End If
         End If
      End If
   Next r_int_NroFil

   Set g_rst_Princi = Nothing
     
   'IMPRIMIR
   crp_Imprim.SelectionFormula = ""
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".RPT_TABLA_TEMP"
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_USUCRE} = '" & modgen_g_str_CodUsu & "' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_TERCRE} = '" & modgen_g_str_NombPC & "' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_TABLA_TEMP.RPT_NOMBRE} = 'REPORTE CONSTITUCION HIPOTECAS' "
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ope_forcof_14.rpt"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Activa(False)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Public Sub gs_IngOper(ByVal p_NumOpe As String)
Dim r_int_Contad     As Integer
Dim r_int_RepNum     As Integer

r_int_RepNum = 0

   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 7) = p_NumOpe Then
         r_int_RepNum = r_int_RepNum + 1
      End If
   Next r_int_Contad
   
   If r_int_RepNum > 0 Then
      MsgBox "El Cliente ya se encuentra ingresado. ", vbExclamation, modgen_g_con_OpeTra
      Exit Sub
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT (TRIM(B.DATGEN_APEPAT) || ' ' || TRIM(B.DATGEN_APEMAT) || ' ' ||  TRIM(B.DATGEN_NOMBRE)) AS CLIENTE,  "
      g_str_Parame = g_str_Parame & "        CASE WHEN INSTR(TRIM(D.PARDES_DESCRI),'EXTRANJERIA') > 0 THEN 'CE' ELSE TRIM(D.PARDES_DESCRI) END || '-' || TRIM(HIPMAE_NDOCLI) AS DOCIDE, "
      g_str_Parame = g_str_Parame & "        CASE WHEN INSTR(TRIM(A.HIPMAE_OPEMVI),'-') > 0 THEN SUBSTR(TRIM(A.HIPMAE_OPEMVI),1,INSTR(TRIM(A.HIPMAE_OPEMVI),'-') -1) "
      g_str_Parame = g_str_Parame & "        ELSE CASE WHEN INSTR(TRIM(A.HIPMAE_OPEMVI),'/') > 0 THEN "
      g_str_Parame = g_str_Parame & "                  SUBSTR(TRIM(A.HIPMAE_OPEMVI),1,INSTR(TRIM(A.HIPMAE_OPEMVI),'/') -1) "
      g_str_Parame = g_str_Parame & "             ELSE TRIM(A.HIPMAE_OPEMVI) END END AS CODFMV, "
      g_str_Parame = g_str_Parame & "        CASE WHEN TRIM(HIPGAR_PARFIC) IS NOT NULL THEN TRIM(HIPGAR_PARFIC) ELSE TRIM(EVALEG_NUMPAR_INM) END AS PARTIDA, "
      g_str_Parame = g_str_Parame & "        CASE WHEN PRODUC_CODIGO='001' THEN 'CRC-PBP' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='002' THEN 'MICASITA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='003' THEN 'CME' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='004' THEN 'MIHOGAR' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='006' THEN 'MICASITA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='007' THEN 'MIVIVIENDA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='009' THEN 'MIVIVIENDA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='010' THEN 'MIVIVIENDA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='011' THEN 'MICASITA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='013' THEN 'MIVIVIENDA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='014' THEN 'MIVIVIENDA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='015' THEN 'MIVIVIENDA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='016' THEN 'MIVIVIENDA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='017' THEN 'MIVIVIENDA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='018' THEN 'MIVIVIENDA' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='019' THEN 'MICASA MAS' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='021' THEN 'MIVIVIENDA MAS BBP' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='022' THEN 'BBP COMPLEMENTO INICIAL' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='023' THEN 'PBP COMPLEMENTO INICIAL' "
      g_str_Parame = g_str_Parame & "             WHEN PRODUC_CODIGO='024' THEN 'MIVIVIENDA TECHO PROPIO' "
      g_str_Parame = g_str_Parame & "        END AS PRODUCTO, "
      g_str_Parame = g_str_Parame & "        TRIM(TRIM(G.DATGEN_APEPAT) || ' ' || TRIM(G.DATGEN_APEMAT) || ' ' ||  TRIM(G.DATGEN_NOMBRE)) AS CONYUGE, "
      g_str_Parame = g_str_Parame & "        CASE WHEN INSTR(TRIM(H.PARDES_DESCRI),'EXTRANJERIA') > 0 THEN 'CE' ELSE TRIM(H.PARDES_DESCRI) END || '-' || TRIM(G.DATGEN_NUMDOC) AS DOCIDE_CONYUGE "
      g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE A "
      g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.HIPMAE_TDOCLI AND B.DATGEN_NUMDOC = HIPMAE_NDOCLI "
      g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.HIPMAE_CODPRD "
      g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 230 AND D.PARDES_CODITE = B.DATGEN_TIPDOC "
      g_str_Parame = g_str_Parame & "         LEFT JOIN TRA_EVALEG E ON E.EVALEG_NUMSOL = A.HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & "         LEFT JOIN CRE_HIPGAR F ON F.HIPGAR_NUMOPE = A.HIPMAE_NUMOPE "
      g_str_Parame = g_str_Parame & "         LEFT JOIN CLI_DATGEN G ON G.DATGEN_TIPDOC = A.HIPMAE_TDOCYG AND G.DATGEN_NUMDOC = A.HIPMAE_NDOCYG "
      g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES H ON H.PARDES_CODGRP = 230 AND H.PARDES_CODITE = G.DATGEN_TIPDOC "
      g_str_Parame = g_str_Parame & "  WHERE A.HIPMAE_NUMOPE = '" & p_NumOpe & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         MsgBox "No existe ninguna Operación registrada con ese Número. ", vbExclamation, modgen_g_con_OpeTra
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         Exit Sub
      Else
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = grd_Listad.Row + 1
         
         grd_Listad.Col = 1
         grd_Listad.Text = g_rst_Princi!CLIENTE
      
         grd_Listad.Col = 2
         grd_Listad.Text = g_rst_Princi!DOCIDE
         
         grd_Listad.Col = 3
         grd_Listad.Text = IIf(IsNull(g_rst_Princi!CODFMV), "", g_rst_Princi!CODFMV)
            
         grd_Listad.Col = 5
         grd_Listad.Text = g_rst_Princi!PRODUCTO
         
         grd_Listad.Col = 7
         grd_Listad.Text = p_NumOpe
               
         grd_Listad.Col = 4
         Do Until g_rst_Princi.EOF
            grd_Listad.Text = grd_Listad.Text & " / " & IIf(IsNull(g_rst_Princi!PARTIDA), "", g_rst_Princi!PARTIDA)
            g_rst_Princi.MoveNext
         Loop
         If InStr(grd_Listad.Text, " / ") = 1 Then grd_Listad.Text = Trim(Mid(grd_Listad.Text, 3))
         Call gs_RefrescaGrid(grd_Listad)
         
      End If
      
      If grd_Listad.Rows > 0 Then
         fs_Activa (True)
         cmd_BusCli.Enabled = True
      End If
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(False)
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
  
   Screen.MousePointer = 0
End Sub

Private Sub fs_Activa(ByVal estado As Boolean)
   cmd_Imprim.Enabled = estado
   cmd_BusCli.Enabled = Not estado
   cmd_Borrar.Enabled = estado
   cmd_Limpia.Enabled = estado
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 660  'ITEM
   grd_Listad.ColWidth(1) = 3950 'CLIENTE
   grd_Listad.ColWidth(2) = 1680 'DNI
   grd_Listad.ColWidth(3) = 2150 'CODIGO FMV
   grd_Listad.ColWidth(4) = 1900 'PARTIDA REGISTRAL
   grd_Listad.ColWidth(5) = 2770 'PRODUCTO
   grd_Listad.ColWidth(6) = 1400 'SELECCIONAR
   grd_Listad.ColWidth(7) = 0    'OPERACION
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter

End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      If grd_Listad.TextMatrix(grd_Listad.Row, 6) = "X" Then
         grd_Listad.TextMatrix(grd_Listad.Row, 6) = ""
      Else
         grd_Listad.TextMatrix(grd_Listad.Row, 6) = "X"
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub
Private Sub pnl_Tit_Item_Click()
   If Len(Trim(pnl_Tit_Item.Tag)) = 0 Or pnl_Tit_Item.Tag = "D" Then
      pnl_Tit_Item.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "N")
   Else
      pnl_Tit_Item.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "N-")
   End If
End Sub
Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub
Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub
Private Sub pnl_Tit_CodFMV_Click()
   If Len(Trim(pnl_Tit_CodFMV.Tag)) = 0 Or pnl_Tit_CodFMV.Tag = "D" Then
      pnl_Tit_CodFMV.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "N")
   Else
      pnl_Tit_CodFMV.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "N-")
   End If
End Sub
Private Sub pnl_Tit_ParReg_Click()
   If Len(Trim(pnl_Tit_ParReg.Tag)) = 0 Or pnl_Tit_ParReg.Tag = "D" Then
      pnl_Tit_ParReg.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "N")
   Else
      pnl_Tit_ParReg.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "N-")
   End If
End Sub
Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub
