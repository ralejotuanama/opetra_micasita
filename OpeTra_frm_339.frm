VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Tas_ActReg_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13620
   Icon            =   "OpeTra_frm_339.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9240
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13635
      _Version        =   65536
      _ExtentX        =   24051
      _ExtentY        =   16298
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   8
         Top             =   810
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_339.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12915
            Picture         =   "OpeTra_frm_339.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_339.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_339.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   900
         Left            =   60
         TabIndex        =   9
         Top             =   1500
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
         _ExtentY        =   1587
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1620
            TabIndex        =   1
            Top             =   465
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            MinValue        =   "0"
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
         Begin VB.Label Label3 
            Caption         =   "Año de Proceso:"
            Height          =   285
            Left            =   150
            TabIndex        =   11
            Top             =   510
            Width           =   1245
         End
         Begin VB.Label Label4 
            Caption         =   "Mes de Proceso:"
            Height          =   315
            Left            =   150
            TabIndex        =   10
            Top             =   150
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   660
            TabIndex        =   13
            Top             =   60
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   660
            TabIndex        =   14
            Top             =   360
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Consulta de Proceso de Actualización de Garantías"
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
            Picture         =   "OpeTra_frm_339.frx":0D6C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6735
         Left            =   60
         TabIndex        =   15
         Top             =   2445
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
         _ExtentY        =   11880
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
            Height          =   6630
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   13410
            _ExtentX        =   23654
            _ExtentY        =   11695
            _Version        =   393216
            Rows            =   31
            Cols            =   14
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_Tas_ActReg_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function ff_ObtieneDatos(ByVal p_MesProc As String, ByVal p_AnioProc As String) As String
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String

   
   If p_MesProc = "12" Then
      r_str_FecIni = Format(CInt(p_AnioProc) + 1, "0000") & "01" & "01"
      r_str_FecFin = Format(CInt(p_AnioProc) + 1, "0000") & "03" & "31"
   Else
      r_str_FecIni = p_AnioProc & Format(CInt(p_MesProc) + 1, "00") & "01"
      r_str_FecFin = p_AnioProc & Format(CInt(p_MesProc) + 3, "00") & "31"
   End If
   
   ff_ObtieneDatos = ""
   ff_ObtieneDatos = ff_ObtieneDatos & "SELECT ACTGAR_NUMOPE AS OPERACION, ACTGAR_FECPRO AS FECHA_PROCESO, ACTGAR_SALCAP AS SALDO_CAPITAL, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       CASE WHEN C.HIPMAE_CODPRD IN (" & moddat_g_str_AgrTMIC & ") THEN 'MICASITA' ELSE 'MIVIVIENDA' END AS PRODUCTO, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       CASE WHEN C.HIPMAE_MONEDA = 1 THEN 'SOLES' ELSE 'DOLARES AMERICANOS' END AS TIPO_MONEDA, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       ACTGAR_TIPDOC ||'-'|| ACTGAR_NUMDOC AS TIPO_DOCUMENTO, TRIM(E.PARDES_DESCRI) AS DISTRITO, C.HIPMAE_FECDES AS DESEMBOLSO, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       TRIM(DATGEN_APEPAT) ||' '|| TRIM(DATGEN_APEMAT) ||' '|| TRIM(DATGEN_NOMBRE) AS CLIENTE, TRIM(H.PARDES_DESCRI) AS SITUACION, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       TRIM(NVL(F.DATGEN_TITULO, ' ')) AS PROYECTO, TRIM(NVL(G.DATGEN_RAZSOC, ' ')) AS PROMOTOR, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       ACTGAR_FECTAS AS FECHA_TASACION, ACTGAR_MESTAS AS MES_TASACION, ACTGAR_ANOTAS AS ANIO_TASACION, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       ACTGAR_ARECON AS AREA_CONSTRUIDA, ACTGAR_ARETER AS AREA_TERRENO, ACTGAR_VALCOM AS VALOR_COMERCIAL, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       ACTGAR_ANOCON AS ANIO_CONSTRUCCION, ACTGAR_MATCON AS MATERIAL_CONSTRUCCION, ACTGAR_CORTER AS CORREGIDO_TERRENO, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       ACTGAR_ANOANT AS ANTIGUEDAD_ACTUAL, ACTGAR_PORDEP AS DEPRECIACION, ACTGAR_VALTER AS VALOR_M2_TERRENO, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       ACTGAR_VALCON AS VALOR_M2_CONSTRUCCION, ACTGAR_VALACT AS VALOR_ACTUAL, ACTGAR_VALACZ AS VALOR_ACTUALIZADO, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       ACTGAR_ESTCON AS ESTADO_CONSERVACION, ACTGAR_VALACC AS VALOR_ACTZ_COM, ACTGAR_CORVAC AS CORREGIDO_ACTUAL, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       ACTGAR_CORCON AS CORREGIDO_CONSTRUCCION, ACTGAR_CORACC AS CORREGIDO_ACT_COM, ACTGAR_PORLTV AS PORCENTAJE_LTV "
   ff_ObtieneDatos = ff_ObtieneDatos & "  FROM CRE_ACTGAR A "
   ff_ObtieneDatos = ff_ObtieneDatos & " INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.ACTGAR_TIPDOC AND TRIM(B.DATGEN_NUMDOC) = TRIM(A.ACTGAR_NUMDOC) "
   ff_ObtieneDatos = ff_ObtieneDatos & " INNER JOIN CRE_HIPMAE C ON C.HIPMAE_NUMOPE = A.ACTGAR_NUMOPE "
   ff_ObtieneDatos = ff_ObtieneDatos & " INNER JOIN CRE_SOLINM D ON D.SOLINM_NUMSOL = C.HIPMAE_NUMSOL "
   ff_ObtieneDatos = ff_ObtieneDatos & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = '101' AND E.PARDES_CODITE = D.SOLINM_UBIGEO "
   ff_ObtieneDatos = ff_ObtieneDatos & "  LEFT JOIN PRY_DATGEN F ON F.DATGEN_CODIGO = C.HIPMAE_PRYINM "
   ff_ObtieneDatos = ff_ObtieneDatos & "  LEFT JOIN EMP_DATGEN G ON G.DATGEN_EMPTDO = F.DATGEN_VENTDO AND G.DATGEN_EMPNDO = F.DATGEN_VENNDO "
   ff_ObtieneDatos = ff_ObtieneDatos & " INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = '027' AND H.PARDES_CODITE = C.HIPMAE_SITUAC "
   ff_ObtieneDatos = ff_ObtieneDatos & " WHERE ACTGAR_MESPRO = '" & p_MesProc & "' "
   ff_ObtieneDatos = ff_ObtieneDatos & "   AND ACTGAR_ANOPRO = '" & p_AnioProc & "' "
   'ff_ObtieneDatos = ff_ObtieneDatos & " ORDER BY DISTRITO, ACTGAR_NUMOPE "
   
   ff_ObtieneDatos = ff_ObtieneDatos & " UNION "
   
   ff_ObtieneDatos = ff_ObtieneDatos & "SELECT HIPMAE_NUMOPE AS OPERACION, TO_CHAR(EVATAS_FECREG) AS FECHA_PROCESO, 0 AS SALDO_CAPITAL, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       CASE WHEN C.HIPMAE_CODPRD IN ('002','006','011') THEN 'MICASITA' ELSE 'MIVIVIENDA' END AS PRODUCTO, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       CASE WHEN C.HIPMAE_MONEDA = 1 THEN 'SOLES' ELSE 'DOLARES AMERICANOS' END AS TIPO_MONEDA, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       TRIM(D.DATGEN_TIPDOC)||'-'||TRIM(D.DATGEN_NUMDOC) AS TIPO_DOCUMENTO, '-' AS DISTRITO, C.HIPMAE_FECDES AS DESEMBOLSO, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       TRIM(D.DATGEN_APEPAT)||' '||TRIM(D.DATGEN_APEMAT)||' '||TRIM(D.DATGEN_NOMBRE) AS CLIENTE, TRIM(H.PARDES_DESCRI) AS SITUACION, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       TRIM(NVL(F.DATGEN_TITULO, ' ')) AS PROYECTO, TRIM(NVL(G.DATGEN_RAZSOC, ' ')) AS PROMOTOR, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       TO_CHAR(A.EVATAS_FECEVA) AS FECHA_TASACION, SUBSTR(A.EVATAS_FECEVA,5,2) AS MES_TASACION, SUBSTR(A.EVATAS_FECEVA,1,4) AS ANIO_TASACION, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       0 AS AREA_CONSTRUIDA,         0 AS AREA_TERRENO,            0 AS VALOR_COMERCIAL, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       '0' AS ANIO_CONSTRUCCION,     '-' AS MATERIAL_CONSTRUCCION, 0 AS CORREGIDO_TERRENO ,"
   ff_ObtieneDatos = ff_ObtieneDatos & "       0 AS ANTIGUEDAD_ACTUAL,       0 AS DEPRECIACION,            0 AS VALOR_M2_TERRENO, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       0 AS VALOR_M2_CONSTRUCCION,   0 AS VALOR_ACTUAL,            0 AS VALOR_ACTUALIZADO, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       '-' AS ESTADO_CONSERVACION,   0 AS VALOR_ACTZ_COM,          0 AS CORREGIDO_ACTUAL, "
   ff_ObtieneDatos = ff_ObtieneDatos & "       0 AS CORREGIDO_CONSTRUCCION,  0 AS CORREGIDO_ACT_COM,       0 AS PORCENTAJE_LTV "
   ff_ObtieneDatos = ff_ObtieneDatos & "  FROM HIS_EVATAS A"
   ff_ObtieneDatos = ff_ObtieneDatos & " INNER JOIN CRE_HIPMAE C ON C.HIPMAE_NUMSOL = A.EVATAS_NUMSOL"
   ff_ObtieneDatos = ff_ObtieneDatos & " INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = C.HIPMAE_TDOCLI AND TRIM(D.DATGEN_NUMDOC) = TRIM(C.HIPMAE_NDOCLI)"
   ff_ObtieneDatos = ff_ObtieneDatos & "  LEFT JOIN PRY_DATGEN F ON F.DATGEN_CODIGO = C.HIPMAE_PRYINM"
   ff_ObtieneDatos = ff_ObtieneDatos & "  LEFT JOIN EMP_DATGEN G ON G.DATGEN_EMPTDO = F.DATGEN_VENTDO AND G.DATGEN_EMPNDO = F.DATGEN_VENNDO"
   ff_ObtieneDatos = ff_ObtieneDatos & " INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = '027' AND H.PARDES_CODITE = C.HIPMAE_SITUAC"
   ff_ObtieneDatos = ff_ObtieneDatos & " WHERE EVATAS_FECREG >= " & r_str_FecIni
   ff_ObtieneDatos = ff_ObtieneDatos & "   AND EVATAS_FECREG <= " & r_str_FecFin
End Function

Private Sub fs_Buscar()
Dim r_str_Param1     As String
Dim r_int_Contad     As Integer
   
   'Obtiene datos de Garantias
   r_str_Param1 = ff_ObtieneDatos(fs_NumeroMes(cmb_PerMes.Text), ipp_PerAno.Text)
   
   If Not gf_EjecutaSQL(r_str_Param1, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Call fs_Activa(True)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   
   'Primera Linea
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.RowHeight(0) = 440
   grd_Listad.WordWrap = True
   
   grd_Listad.Col = 0:   grd_Listad.Text = "ITEM":                   grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 1:   grd_Listad.Text = "OPERACION":              grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 2:   grd_Listad.Text = "PRODUCTO":               grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 3:   grd_Listad.Text = "FECHA PROCESO":          grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 4:   grd_Listad.Text = "TIPO DOCUMENTO":         grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 5:   grd_Listad.Text = "NOMBRE DEL CLIENTE":     grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 6:   grd_Listad.Text = "NOMBRE DEL PROYECTO":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 7:   grd_Listad.Text = "NOMBRE DEL PROMOTOR":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 8:   grd_Listad.Text = "FECHA DESEMBOLSO":       grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 9:   grd_Listad.Text = "TIPO MONEDA":            grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 10:  grd_Listad.Text = "SALDO CAPITAL":          grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 11:  grd_Listad.Text = "DISTRITO":               grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 12:  grd_Listad.Text = "FECHA TASACION":         grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 13:  grd_Listad.Text = "MES TASACION":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 14:  grd_Listad.Text = "AÑO TASACION":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 15:  grd_Listad.Text = "AREA CONSTRUIDA":        grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 16:  grd_Listad.Text = "AREA TERRENO":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 17:  grd_Listad.Text = "VALOR COMERCIAL":        grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 18:  grd_Listad.Text = "AÑO CONSTRUCCION":       grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 19:  grd_Listad.Text = "MATERIAL CONSTRUCCION":  grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 20:  grd_Listad.Text = "ESTADO CONSERVACION":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 21:  grd_Listad.Text = "ANTIGUEDAD (AÑOS)":      grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 22:  grd_Listad.Text = "DEPRECIACION (%)":       grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 23:  grd_Listad.Text = "VALOR M2 TERRENO":       grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 24:  grd_Listad.Text = "VALOR M2 CONSTRUCCION":  grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 25:  grd_Listad.Text = "VALOR ACTUAL":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 26:  grd_Listad.Text = "VALOR ACTUALIZADO":      grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 27:  grd_Listad.Text = "VALACT/VALCOM (%)":      grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 28:  grd_Listad.Text = "CORREGIDO VAL TER":      grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 29:  grd_Listad.Text = "CORREGIDO VAL CONST":    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 30:  grd_Listad.Text = "CORREGIDO VAL ACTUAL":   grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 31:  grd_Listad.Text = "CORREGIDO ACT/COM (%)":  grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 32:  grd_Listad.Text = "LTV (%)":                grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 33:  grd_Listad.Text = "SITUACION":              grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 34:  grd_Listad.Text = "COMENTARIO":             grd_Listad.CellAlignment = flexAlignCenterCenter
   
   grd_Listad.Redraw = False
   r_int_Contad = 0
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      r_int_Contad = r_int_Contad + 1
      
      'Numero de Item
      grd_Listad.Col = 0: grd_Listad.Text = Format(r_int_Contad, "0000")
      
      'Numero Operacion
      grd_Listad.Col = 1: grd_Listad.Text = gf_Formato_NumOpe(Trim(g_rst_Princi!OPERACION & ""))
      
      'Producto
      grd_Listad.Col = 2: grd_Listad.Text = Trim(g_rst_Princi!PRODUCTO & "")
      
      'Fecha de Proceso
      grd_Listad.Col = 3: grd_Listad.Text = Right(g_rst_Princi!FECHA_PROCESO, 2) & "/" & Mid(g_rst_Princi!FECHA_PROCESO, 5, 2) & "/" & Left(g_rst_Princi!FECHA_PROCESO, 4)
      
      'Tipo y Nro.Documento
      grd_Listad.Col = 4: grd_Listad.Text = g_rst_Princi!TIPO_DOCUMENTO
      
      'Nombre del Cliente
      grd_Listad.Col = 5: grd_Listad.Text = g_rst_Princi!CLIENTE
      
      'Nombre del proyecto
      If Not IsNull(g_rst_Princi!PROYECTO) Then
         grd_Listad.Col = 6: grd_Listad.Text = g_rst_Princi!PROYECTO
      End If
      
      'Nombre del promotor
      If Not IsNull(g_rst_Princi!PROMOTOR) Then
         grd_Listad.Col = 7: grd_Listad.Text = g_rst_Princi!PROMOTOR
      End If
      
      'Fecha de Desembolso
      grd_Listad.Col = 8: grd_Listad.Text = Right(g_rst_Princi!DESEMBOLSO, 2) & "/" & Mid(g_rst_Princi!DESEMBOLSO, 5, 2) & "/" & Left(g_rst_Princi!DESEMBOLSO, 4)
      
      'Tipo de Moneda
      grd_Listad.Col = 9: grd_Listad.Text = g_rst_Princi!TIPO_MONEDA
      
      'Saldo Capital
      grd_Listad.Col = 10: grd_Listad.Text = Format(g_rst_Princi!SALDO_CAPITAL, "###,##0.00")
      
      'Distrito
      grd_Listad.Col = 11: grd_Listad.Text = Trim(g_rst_Princi!DISTRITO)
      
      'Fecha Tasacion
      grd_Listad.Col = 12: grd_Listad.Text = Right(g_rst_Princi!FECHA_TASACION, 2) & "/" & Mid(g_rst_Princi!FECHA_TASACION, 5, 2) & "/" & Left(g_rst_Princi!FECHA_TASACION, 4)
      
      'Mes Tasacion
      grd_Listad.Col = 13: grd_Listad.Text = g_rst_Princi!MES_TASACION
      
      'Año Tasacion
      grd_Listad.Col = 14: grd_Listad.Text = g_rst_Princi!ANIO_TASACION
      
      'Area Construida
      grd_Listad.Col = 15: grd_Listad.Text = Format(g_rst_Princi!AREA_CONSTRUIDA, "###,###,##0.00")
      
      'Area Terreno
      grd_Listad.Col = 16: grd_Listad.Text = Format(g_rst_Princi!AREA_TERRENO, "###,###,##0.00")
      
      'Valor Comercial
      grd_Listad.Col = 17: grd_Listad.Text = Format(g_rst_Princi!VALOR_COMERCIAL, "###,###,##0.00")
      
      'Año Construccion
      grd_Listad.Col = 18: grd_Listad.Text = g_rst_Princi!ANIO_CONSTRUCCION
      
      'Material Construccion
      grd_Listad.Col = 19: grd_Listad.Text = g_rst_Princi!MATERIAL_CONSTRUCCION
      
      'Estado Conservacion
      grd_Listad.Col = 20: grd_Listad.Text = g_rst_Princi!ESTADO_CONSERVACION
      
      'Antiguedad
      grd_Listad.Col = 21: grd_Listad.Text = g_rst_Princi!ANTIGUEDAD_ACTUAL
      
      'Depreciacion
      grd_Listad.Col = 22: grd_Listad.Text = g_rst_Princi!DEPRECIACION
      
      'Valor m2 Terreno
      grd_Listad.Col = 23: grd_Listad.Text = Format(g_rst_Princi!VALOR_M2_TERRENO, "###,###,##0.00")
      
      'Valor m2 Construccion
      grd_Listad.Col = 24: grd_Listad.Text = Format(g_rst_Princi!VALOR_M2_CONSTRUCCION, "###,###,##0.00")
      
      'Valor Actual
      grd_Listad.Col = 25: grd_Listad.Text = Format(g_rst_Princi!VALOR_ACTUAL, "###,###,##0.00")
      
      'Valor Actualizado
      grd_Listad.Col = 26: grd_Listad.Text = Format(g_rst_Princi!VALOR_ACTUALIZADO, "###,###,##0.00")
      
      'Relacion ValAct / ValCom
      grd_Listad.Col = 27: grd_Listad.Text = Format(g_rst_Princi!VALOR_ACTZ_COM, "###,###,##0.00")
      
      'Valor Corregido - Valor Terreno
      grd_Listad.Col = 28: grd_Listad.Text = Format(g_rst_Princi!CORREGIDO_TERRENO, "###,###,##0.00")
      
      'Valor Corregido - Valor Construccion
      grd_Listad.Col = 29: grd_Listad.Text = Format(g_rst_Princi!CORREGIDO_CONSTRUCCION, "###,###,##0.00")
      
      'Valor Corregido - Valor Actualizado
      grd_Listad.Col = 30: grd_Listad.Text = Format(g_rst_Princi!CORREGIDO_ACTUAL, "###,###,##0.00")
      
      'Valor Corregido - Relacion ValAct / ValCom
      grd_Listad.Col = 31: grd_Listad.Text = Format(g_rst_Princi!CORREGIDO_ACT_COM, "###,###,##0.00")
      
      'LTV
      grd_Listad.Col = 32: grd_Listad.Text = Format(g_rst_Princi!PORCENTAJE_LTV, "##0.00")
      
      'Situacion
      grd_Listad.Col = 33: grd_Listad.Text = Format(g_rst_Princi!SITUACION, "##0.00")
      
      'Comentario
      grd_Listad.Col = 34: grd_Listad.Text = fs_Obtiene_Tasacion(Trim(g_rst_Princi!OPERACION), fs_NumeroMes(cmb_PerMes.Text), ipp_PerAno.Text)
      
      g_rst_Princi.MoveNext
   Loop
   
   With grd_Listad
      .FixedCols = 1
      .FixedRows = 1
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Function fs_Obtiene_Tasacion(ByVal p_NumOpe As String, ByVal p_PerMes As String, ByVal p_PerAno As String) As String
Dim r_str_Param2     As String
Dim r_rst_Cadena     As adodb.Recordset
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String

   fs_Obtiene_Tasacion = "OPERACION REVISADA EN EL PROCESO"
   If p_PerMes = "12" Then
      r_str_FecIni = Format(CInt(p_PerAno) + 1, "0000") & "01" & "01"
      r_str_FecFin = Format(CInt(p_PerAno) + 1, "0000") & "03" & "31"
   Else
      r_str_FecIni = p_PerAno & Format(CInt(p_PerMes) + 1, "00") & "01"
      r_str_FecFin = p_PerAno & Format(CInt(p_PerMes) + 3, "00") & "31"
   End If
  
   'Obtiene datos de Garantias
   r_str_Param2 = ""
   r_str_Param2 = r_str_Param2 & "SELECT * "
   r_str_Param2 = r_str_Param2 & "  FROM HIS_EVATAS "
   r_str_Param2 = r_str_Param2 & " WHERE EVATAS_NUMSOL IN (SELECT HIPMAE_NUMSOL FROM CRE_HIPMAE WHERE HIPMAE_NUMOPE = '" & p_NumOpe & "') "
   r_str_Param2 = r_str_Param2 & "   AND EVATAS_FECREG >= " & r_str_FecIni
   r_str_Param2 = r_str_Param2 & "   AND EVATAS_FECREG <= " & r_str_FecFin
   
   If Not gf_EjecutaSQL(r_str_Param2, r_rst_Cadena, 3) Then
      Exit Function
   End If
   
   If r_rst_Cadena.BOF And r_rst_Cadena.EOF Then
      r_rst_Cadena.Close
      Set r_rst_Cadena = Nothing
      Exit Function
   End If
   
   r_rst_Cadena.MoveFirst
   fs_Obtiene_Tasacion = "OPERACION CON TASACION ACTUALIZADA"
   
   r_rst_Cadena.Close
   Set r_rst_Cadena = Nothing
End Function

Private Sub cmd_Buscar_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el mes de proceso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If CInt(ipp_PerAno.Text) < 2012 Then
      MsgBox "Ingrese correctamente el año de proceso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de consultar información?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el mes de proceso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If CInt(ipp_PerAno.Text) < 2012 Then
      MsgBox "Ingrese correctamente el año de proceso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_PerMes)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_CentraForm(Me)
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.Cols = 35
   grd_Listad.ColWidth(0) = 800        'ITEM
   grd_Listad.ColWidth(1) = 1300       'OPERACION
   grd_Listad.ColWidth(2) = 1300       'PRODUCTO
   grd_Listad.ColWidth(3) = 1200       'FECHA PROCESO
   grd_Listad.ColWidth(4) = 1200       'TIPO DOCUMENTO
   grd_Listad.ColWidth(5) = 3500       'NOMBRE DEL CLIENTE
   grd_Listad.ColWidth(6) = 3500       'NOMBRE DEL PROYECTO
   grd_Listad.ColWidth(7) = 3500       'NOMBRE DEL PROMOTOR
   grd_Listad.ColWidth(8) = 1200       'FECHA DESEMBOLSO
   grd_Listad.ColWidth(9) = 2100       'TIPO MONEDA
   grd_Listad.ColWidth(10) = 1300       'SALDO CAPITAL
   grd_Listad.ColWidth(11) = 2500      'DISTRITO
   grd_Listad.ColWidth(12) = 1200      'FECHA TASACION
   grd_Listad.ColWidth(13) = 1100      'MES TASACION
   grd_Listad.ColWidth(14) = 1100      'AÑO TASACION
   grd_Listad.ColWidth(15) = 1200      'AREA CONTRUIDA
   grd_Listad.ColWidth(16) = 1200      'AREA TERRENO
   grd_Listad.ColWidth(17) = 1200      'VALOR COMERCIAL
   grd_Listad.ColWidth(18) = 1500      'AÑO CONSTRUCCION
   grd_Listad.ColWidth(19) = 1500      'MATERIAL CONSTRUCCION
   grd_Listad.ColWidth(20) = 1500      'ESTADO CONSERVACION
   grd_Listad.ColWidth(21) = 1400      'ANTIGUEDAD
   grd_Listad.ColWidth(22) = 1500      'DEPRECIACION
   grd_Listad.ColWidth(23) = 1500      'VALOR M2 TERRENO
   grd_Listad.ColWidth(24) = 1500      'VALOR M2 CONSTRUCCION
   grd_Listad.ColWidth(25) = 1500      'VALOR ACTUAL
   grd_Listad.ColWidth(26) = 1500      'VALOR ACTUALIZADO
   grd_Listad.ColWidth(27) = 1500      'RELACION VAL ACT / VAL COM
   grd_Listad.ColWidth(28) = 1500      'VALOR CORREGIDO - VALOR TERRENO
   grd_Listad.ColWidth(29) = 1500      'VALOR CORREGIDO - VALOR CONSTRUCCION
   grd_Listad.ColWidth(30) = 1500      'VALOR CORREGIDO - VALOR ACTUALIZADO
   grd_Listad.ColWidth(31) = 1600      'VALOR CORREGIDO - RELACION VAL ACT / VAL COM
   grd_Listad.ColWidth(32) = 1200      'LTV
   grd_Listad.ColWidth(33) = 1500      'SITUACION
   grd_Listad.ColWidth(34) = 4000      'COMENTARIO
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignLeftCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
   grd_Listad.ColAlignment(12) = flexAlignCenterCenter
   grd_Listad.ColAlignment(13) = flexAlignCenterCenter
   grd_Listad.ColAlignment(14) = flexAlignCenterCenter
   grd_Listad.ColAlignment(15) = flexAlignRightCenter
   grd_Listad.ColAlignment(16) = flexAlignRightCenter
   grd_Listad.ColAlignment(17) = flexAlignRightCenter
   grd_Listad.ColAlignment(18) = flexAlignCenterCenter
   grd_Listad.ColAlignment(19) = flexAlignCenterCenter
   grd_Listad.ColAlignment(20) = flexAlignCenterCenter
   grd_Listad.ColAlignment(21) = flexAlignCenterCenter
   grd_Listad.ColAlignment(22) = flexAlignRightCenter
   grd_Listad.ColAlignment(23) = flexAlignRightCenter
   grd_Listad.ColAlignment(24) = flexAlignRightCenter
   grd_Listad.ColAlignment(25) = flexAlignRightCenter
   grd_Listad.ColAlignment(26) = flexAlignRightCenter
   grd_Listad.ColAlignment(27) = flexAlignRightCenter
   grd_Listad.ColAlignment(28) = flexAlignRightCenter
   grd_Listad.ColAlignment(29) = flexAlignRightCenter
   grd_Listad.ColAlignment(30) = flexAlignRightCenter
   grd_Listad.ColAlignment(31) = flexAlignRightCenter
   grd_Listad.ColAlignment(32) = flexAlignRightCenter
   grd_Listad.ColAlignment(33) = flexAlignCenterCenter
   grd_Listad.ColAlignment(34) = flexAlignLeftCenter
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
End Sub

Private Sub fs_Limpia()
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = Year(date)
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Function fs_NumeroMes(mes As String) As String
   Select Case mes
      Case "ENERO":     fs_NumeroMes = "01"
      Case "FEBRERO":   fs_NumeroMes = "02"
      Case "MARZO":     fs_NumeroMes = "03"
      Case "ABRIL":     fs_NumeroMes = "04"
      Case "MAYO":      fs_NumeroMes = "05"
      Case "JUNIO":     fs_NumeroMes = "06"
      Case "JULIO":     fs_NumeroMes = "07"
      Case "AGOSTO":    fs_NumeroMes = "08"
      Case "SETIEMBRE": fs_NumeroMes = "09"
      Case "OCTUBRE":   fs_NumeroMes = "10"
      Case "NOVIEMBRE": fs_NumeroMes = "11"
      Case "DICIEMBRE": fs_NumeroMes = "12"
   End Select
End Function

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_PerMes.Enabled = p_Activa
   ipp_PerAno.Enabled = p_Activa
   grd_Listad.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_Contad     As Integer
Dim r_int_NroFil     As Integer

   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      'Titulo
      .Cells(2, 1) = "REPORTE DE CONSULTA DE GARANTIAS - PERIODO : " & Trim(cmb_PerMes.Text) & " / " & Trim(ipp_PerAno.Text)
      .Range(.Cells(2, 1), .Cells(2, 35)).Merge
      .Range("A2:AA2").HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = 4
      .Columns("A").ColumnWidth = 5:    .Columns("A").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 1) = "ITEM"
      .Columns("B").ColumnWidth = 14:   .Columns("B").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 2) = "OPERACION"
      .Columns("C").ColumnWidth = 14:   .Columns("C").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 3) = "PRODUCTO"
      .Columns("D").ColumnWidth = 14:   .Columns("D").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 4) = "FECHA DE PROCESO"
      .Columns("E").ColumnWidth = 15:   .Columns("E").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 5) = "DOCUMENTO IDENTIFICACION"
      .Columns("F").ColumnWidth = 40:   .Columns("F").HorizontalAlignment = xlHAlignLeft:     .Cells(r_int_NroFil, 6) = "NOMBRE DEL CLIENTE"
      .Columns("G").ColumnWidth = 45:   .Columns("G").HorizontalAlignment = xlHAlignLeft:     .Cells(r_int_NroFil, 7) = "NOMBRE DEL PROYECTO"
      .Columns("H").ColumnWidth = 50:   .Columns("F").HorizontalAlignment = xlHAlignLeft:     .Cells(r_int_NroFil, 8) = "NOMBRE DEL PROMOTOR"
      .Columns("I").ColumnWidth = 14:   .Columns("I").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 9) = "FECHA DESEMBOLSO"
      .Columns("J").ColumnWidth = 22:   .Columns("J").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 10) = "TIPO DE MONEDA"
      .Columns("K").ColumnWidth = 16:   .Columns("K").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 11) = "SALDO CAPITAL"
      .Columns("L").ColumnWidth = 22:   .Columns("L").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 12) = "DISTRITO"
      .Columns("M").ColumnWidth = 14:   .Columns("M").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 13) = "FECHA DE TASACION"
      .Columns("N").ColumnWidth = 12:   .Columns("N").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 14) = "MES DE TASACION"
      .Columns("O").ColumnWidth = 12:   .Columns("O").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 15) = "AÑO DE TASACION"
      .Columns("P").ColumnWidth = 12:   .Columns("P").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 16) = "AREA CONSTRUIDA"
      .Columns("Q").ColumnWidth = 12:   .Columns("Q").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 17) = "AREA DEL TERRENO"
      .Columns("R").ColumnWidth = 13:   .Columns("R").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 18) = "VALOR COMERCIAL"
      .Columns("S").ColumnWidth = 15:   .Columns("S").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 19) = "AÑO CONSTRUCCION"
      .Columns("T").ColumnWidth = 15:   .Columns("T").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 20) = "MATERIAL CONSTRUCCION"
      .Columns("U").ColumnWidth = 15:   .Columns("U").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 21) = "ESTADO CONSERVACION"
      .Columns("V").ColumnWidth = 15:   .Columns("V").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 22) = "ANTIGUEDAD (AÑOS)"
      .Columns("W").ColumnWidth = 15:   .Columns("W").HorizontalAlignment = xlHAlignCenter:   .Cells(r_int_NroFil, 23) = "DEPRECIACION (%)"
      .Columns("X").ColumnWidth = 13:   .Columns("X").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 24) = "VALOR M2 TERRENO"
      .Columns("Y").ColumnWidth = 15:   .Columns("Y").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 25) = "VALOR M2 CONSTRUCCION"
      .Columns("Z").ColumnWidth = 14:   .Columns("Z").HorizontalAlignment = xlHAlignRight:    .Cells(r_int_NroFil, 26) = "VALOR ACTUAL"
      .Columns("AA").ColumnWidth = 15:  .Columns("AA").HorizontalAlignment = xlHAlignRight:   .Cells(r_int_NroFil, 27) = "VALOR ACTUALIZADO"
      .Columns("AB").ColumnWidth = 16:  .Columns("AB").HorizontalAlignment = xlHAlignRight:   .Cells(r_int_NroFil, 28) = "VALACT/VALCOM (%)"
      .Columns("AC").ColumnWidth = 15:  .Columns("AC").HorizontalAlignment = xlHAlignRight:   .Cells(r_int_NroFil, 29) = "CORREGIDO VALOR TERRENO"
      .Columns("AD").ColumnWidth = 16:  .Columns("AD").HorizontalAlignment = xlHAlignRight:   .Cells(r_int_NroFil, 30) = "CORREGIDO VALOR CONSTR"
      .Columns("AE").ColumnWidth = 16:  .Columns("AE").HorizontalAlignment = xlHAlignRight:   .Cells(r_int_NroFil, 31) = "CORREGIDO VALOR ACTUALIZ"
      .Columns("AF").ColumnWidth = 16:  .Columns("AF").HorizontalAlignment = xlHAlignRight:   .Cells(r_int_NroFil, 32) = "CORREGIDO ACT/COM (%)"
      .Columns("AG").ColumnWidth = 16:  .Columns("AG").HorizontalAlignment = xlHAlignRight:   .Cells(r_int_NroFil, 33) = "LTV (%)"
      .Columns("AH").ColumnWidth = 16:  .Columns("AH").HorizontalAlignment = xlHAlignCenter:  .Cells(r_int_NroFil, 34) = "SITUACION"
      .Columns("AI").ColumnWidth = 50:  .Columns("AI").HorizontalAlignment = xlHAlignCenter:  .Cells(r_int_NroFil, 35) = "COMENTARIO"
      
      .Range(.Cells(1, 1), .Cells(4, 35)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 35)).WrapText = True
      .Range(.Cells(4, 1), .Cells(4, 35)).VerticalAlignment = xlCenter
      .Range(.Cells(4, 1), .Cells(4, 35)).HorizontalAlignment = xlCenter
      .Range(.Cells(4, 1), .Cells(4, 35)).Interior.Color = RGB(146, 208, 80)
      
      For r_int_Contad = 1 To grd_Listad.Rows - 1
         r_int_NroFil = r_int_NroFil + 1
         .Cells(r_int_NroFil, 1) = grd_Listad.TextMatrix(r_int_Contad, 0)
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_Contad, 1)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_Contad, 2)
         .Cells(r_int_NroFil, 4) = "'" & grd_Listad.TextMatrix(r_int_Contad, 3)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_Contad, 4)
         .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_Contad, 5)
         .Cells(r_int_NroFil, 7) = grd_Listad.TextMatrix(r_int_Contad, 6)
         .Cells(r_int_NroFil, 8) = grd_Listad.TextMatrix(r_int_Contad, 7)
         .Cells(r_int_NroFil, 9) = grd_Listad.TextMatrix(r_int_Contad, 8)
         .Cells(r_int_NroFil, 10) = grd_Listad.TextMatrix(r_int_Contad, 9)
         .Cells(r_int_NroFil, 11) = "'" & grd_Listad.TextMatrix(r_int_Contad, 10)
         .Cells(r_int_NroFil, 12) = grd_Listad.TextMatrix(r_int_Contad, 11)
         .Cells(r_int_NroFil, 13) = grd_Listad.TextMatrix(r_int_Contad, 12)
         .Cells(r_int_NroFil, 14) = grd_Listad.TextMatrix(r_int_Contad, 13)
         .Cells(r_int_NroFil, 15) = grd_Listad.TextMatrix(r_int_Contad, 14)
         .Cells(r_int_NroFil, 16) = grd_Listad.TextMatrix(r_int_Contad, 15)
         .Cells(r_int_NroFil, 17) = grd_Listad.TextMatrix(r_int_Contad, 16)
         .Cells(r_int_NroFil, 18) = grd_Listad.TextMatrix(r_int_Contad, 17)
         .Cells(r_int_NroFil, 19) = grd_Listad.TextMatrix(r_int_Contad, 18)
         .Cells(r_int_NroFil, 20) = grd_Listad.TextMatrix(r_int_Contad, 19)
         .Cells(r_int_NroFil, 21) = grd_Listad.TextMatrix(r_int_Contad, 20)
         .Cells(r_int_NroFil, 22) = grd_Listad.TextMatrix(r_int_Contad, 21)
         .Cells(r_int_NroFil, 23) = grd_Listad.TextMatrix(r_int_Contad, 22)
         .Cells(r_int_NroFil, 24) = grd_Listad.TextMatrix(r_int_Contad, 23)
         .Cells(r_int_NroFil, 25) = grd_Listad.TextMatrix(r_int_Contad, 24)
         .Cells(r_int_NroFil, 26) = grd_Listad.TextMatrix(r_int_Contad, 25)
         .Cells(r_int_NroFil, 27) = grd_Listad.TextMatrix(r_int_Contad, 26)
         .Cells(r_int_NroFil, 28) = grd_Listad.TextMatrix(r_int_Contad, 27)
         .Cells(r_int_NroFil, 29) = grd_Listad.TextMatrix(r_int_Contad, 28)
         .Cells(r_int_NroFil, 30) = grd_Listad.TextMatrix(r_int_Contad, 29)
         .Cells(r_int_NroFil, 31) = grd_Listad.TextMatrix(r_int_Contad, 30)
         .Cells(r_int_NroFil, 32) = grd_Listad.TextMatrix(r_int_Contad, 31)
         .Cells(r_int_NroFil, 33) = grd_Listad.TextMatrix(r_int_Contad, 32)
         .Cells(r_int_NroFil, 34) = grd_Listad.TextMatrix(r_int_Contad, 33)
         .Cells(r_int_NroFil, 35) = grd_Listad.TextMatrix(r_int_Contad, 34)
         .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 35)).Borders(xlEdgeTop).LineStyle = xlContinuous
      Next r_int_Contad
   
      .Range(.Cells(4, 1), .Cells(r_int_NroFil, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 2), .Cells(r_int_NroFil, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 3), .Cells(r_int_NroFil, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 4), .Cells(r_int_NroFil, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 5), .Cells(r_int_NroFil, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 6), .Cells(r_int_NroFil, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 7), .Cells(r_int_NroFil, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 8), .Cells(r_int_NroFil, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 9), .Cells(r_int_NroFil, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 10), .Cells(r_int_NroFil, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 11), .Cells(r_int_NroFil, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 12), .Cells(r_int_NroFil, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 13), .Cells(r_int_NroFil, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 14), .Cells(r_int_NroFil, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 15), .Cells(r_int_NroFil, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 16), .Cells(r_int_NroFil, 16)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 17), .Cells(r_int_NroFil, 17)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 18), .Cells(r_int_NroFil, 18)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 19), .Cells(r_int_NroFil, 19)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 20), .Cells(r_int_NroFil, 20)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 21), .Cells(r_int_NroFil, 21)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 22), .Cells(r_int_NroFil, 22)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 23), .Cells(r_int_NroFil, 23)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 24), .Cells(r_int_NroFil, 24)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 25), .Cells(r_int_NroFil, 25)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 26), .Cells(r_int_NroFil, 26)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 27), .Cells(r_int_NroFil, 27)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 28), .Cells(r_int_NroFil, 28)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 29), .Cells(r_int_NroFil, 29)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 30), .Cells(r_int_NroFil, 30)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 31), .Cells(r_int_NroFil, 31)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 32), .Cells(r_int_NroFil, 32)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 33), .Cells(r_int_NroFil, 33)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 34), .Cells(r_int_NroFil, 34)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 35), .Cells(r_int_NroFil, 35)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 36), .Cells(r_int_NroFil, 36)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      .Range(.Cells(4, 1), .Cells(4, 35)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NroFil + 1, 1), .Cells(r_int_NroFil + 1, 35)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Range(.Cells(5, 15), .Cells(r_int_NroFil, 17)).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      .Range(.Cells(5, 23), .Cells(r_int_NroFil, 32)).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub
