VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_RegDes_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15045
   Icon            =   "OpeTra_frm_808.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7630
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   15090
      _Version        =   65536
      _ExtentX        =   26617
      _ExtentY        =   13458
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
         TabIndex        =   11
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
         Begin VB.CommandButton cmd_Aprobar 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_808.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Aprobar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Rechazar 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_808.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Rechazar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Seguimiento 
            Height          =   585
            Left            =   3630
            Picture         =   "OpeTra_frm_808.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Seguimiento por Instancias"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14370
            Picture         =   "OpeTra_frm_808.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_808.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_808.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Buscar Operaciones"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SegSol 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_808.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Detalle de la Operación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3030
            Picture         =   "OpeTra_frm_808.frx":1D82
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel pnl_Filtros 
         Height          =   765
         Left            =   30
         TabIndex        =   12
         Top             =   1440
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
         Begin VB.CheckBox chk_Estado 
            Caption         =   "Todas las Instancias"
            Height          =   315
            Left            =   1110
            TabIndex        =   9
            Top             =   420
            Value           =   1  'Checked
            Width           =   2685
         End
         Begin VB.ComboBox cmb_Estado 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   90
            Width           =   3975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Instancias:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   640
         Left            =   30
         TabIndex        =   14
         Top             =   60
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
         _ExtentY        =   1129
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
            TabIndex        =   15
            Top             =   60
            Width           =   8835
            _Version        =   65536
            _ExtentX        =   15584
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Operaciones a Desembolsar al Promotor"
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
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   13800
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   14385
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_808.frx":208C
            Top             =   100
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   5300
         Left            =   30
         TabIndex        =   16
         Top             =   2250
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
         _ExtentY        =   9349
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
            Height          =   4980
            Left            =   60
            TabIndex        =   17
            Top             =   360
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   8784
            _Version        =   393216
            Rows            =   45
            Cols            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_ConHip 
            Height          =   285
            Left            =   12120
            TabIndex        =   18
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2558
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cons. Hipotecario"
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
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   3510
            TabIndex        =   19
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
            Left            =   4815
            TabIndex        =   20
            Top             =   60
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6209
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
         Begin Threed.SSPanel pnl_Tit_SitAct 
            Height          =   285
            Left            =   9435
            TabIndex        =   21
            Top             =   60
            Width           =   2700
            _Version        =   65536
            _ExtentX        =   4762
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Instancia"
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
         Begin Threed.SSPanel pnl_Tit_FecReg 
            Height          =   285
            Left            =   8310
            TabIndex        =   22
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Registro"
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
            Left            =   90
            TabIndex        =   23
            Top             =   60
            Width           =   3435
            _Version        =   65536
            _ExtentX        =   6059
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
         Begin Threed.SSPanel pnl_Selecc 
            Height          =   285
            Left            =   13560
            TabIndex        =   24
            Top             =   60
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   " Selección"
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
               Left            =   820
               TabIndex        =   25
               Top             =   10
               Width           =   255
            End
         End
      End
   End
End
Attribute VB_Name = "frm_RegDes_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chk_Estado_Click()
   Call Estado_Ctrl
   If chk_Estado.Value = 1 Then
      cmb_Estado.ListIndex = -1
      cmb_Estado.Enabled = False
      Call gs_SetFocus(cmd_Buscar)
   ElseIf chk_Estado.Value = 0 Then
      cmb_Estado.Enabled = True
      Call gs_SetFocus(cmb_Estado)
   End If
End Sub

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If chkSeleccionar.Value = 0 Then
      For r_Fila = 0 To grd_Listad.Rows - 1
          If UCase(grd_Listad.TextMatrix(r_Fila, 20)) = "X" Then
             grd_Listad.TextMatrix(r_Fila, 20) = ""
          End If
      Next r_Fila
   End If
   If chkSeleccionar.Value = 1 Then
      For r_Fila = 0 To grd_Listad.Rows - 1
          If UCase(grd_Listad.TextMatrix(r_Fila, 20)) = "" Then
             grd_Listad.TextMatrix(r_Fila, 20) = "X"
          End If
      Next r_Fila
   End If
   Call gs_RefrescaGrid(grd_Listad)
End Sub

Private Sub cmd_Aprobar_Click()
    'APROBADO
    Call fs_Guardar_Eva(1)
End Sub

Private Sub cmd_Rechazar_Click()
   'RECHAZO
   Call fs_Guardar_Eva(2)
End Sub

Private Sub fs_Guardar_Eva(p_Estado As Integer)
Dim r_str_NumOpe   As String
Dim r_int_NumFil   As Integer
Dim r_int_Estado   As Boolean
Dim r_str_CadAux   As String
Dim r_str_TipEva   As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   r_int_Estado = False
   r_str_NumOpe = ""
   r_str_CadAux = ""
   r_str_TipEva = ""
   
   For r_int_NumFil = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_NumFil, 20) = "X" Then
          r_int_Estado = True
          r_str_NumOpe = r_str_NumOpe & Trim(grd_Listad.TextMatrix(r_int_NumFil, 11)) & "','"
       End If
   Next
   If r_int_Estado = False Then
      r_str_NumOpe = ""
      MsgBox "No hay ninguna fila seleccionada.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   Else
      r_str_NumOpe = "'" & Mid(r_str_NumOpe, 1, Len(r_str_NumOpe) - 2)
   End If
   
   If p_Estado = 1 Then
      If MsgBox("¿Seguro que desea aprobar los registros seleccionados?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Screen.MousePointer = 0
         Exit Sub
      End If
   Else
      If MsgBox("¿Seguro que desea rechazar los registros seleccionados?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Screen.MousePointer = 0
         Exit Sub
      End If
   End If
   
   If p_Estado = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT DESCAB_NUMOPE, HIPMAE_MTOPRE, HIPMAE_MONEDA, CALCULADO - SOLMAE_MTOGCI AS CALCULADO, IMP_TOTDET, REG_TOTDET "
      g_str_Parame = g_str_Parame & "  FROM (SELECT A.DESCAB_NUMOPE, B.HIPMAE_MTOPRE, B.HIPMAE_MONEDA, "
      g_str_Parame = g_str_Parame & "               CASE WHEN HIPMAE_MONEDA = 1 THEN "
      g_str_Parame = g_str_Parame & "                    CASE WHEN B.HIPMAE_CODPRD = '024' THEN "
      g_str_Parame = g_str_Parame & "                              DECODE(HIPMAE_PRYMCS,1, HIPMAE_MTOPRE + SOLMAE_FMVBBP + SOLMAE_PBPMTO + SOLMAE_BMSMTO + SOLMAE_AFPMTO, HIPMAE_MTOPRE + SOLMAE_BMSMTO + SOLMAE_AFPMTO) "
      g_str_Parame = g_str_Parame & "                         WHEN B.HIPMAE_CODPRD <> '019' AND (SELECT INSTR(X.AGRPRO_DESCRI, B.HIPMAE_CODPRD) FROM CRE_AGRPRO X WHERE X.AGRPRO_CODAGR = 'AGR1FMV') > 0 THEN "
      g_str_Parame = g_str_Parame & "                              B.HIPMAE_MTOPRE + C.SOLMAE_FMVBBP + C.SOLMAE_AFPMTO + C.SOLMAE_PBPMTO + C.SOLMAE_BMSMTO "
      g_str_Parame = g_str_Parame & "                         WHEN B.HIPMAE_CODPRD = '011' THEN "
      g_str_Parame = g_str_Parame & "                              HIPMAE_MTOPRE + SOLMAE_AFPMTO "
      g_str_Parame = g_str_Parame & "                    END "
      g_str_Parame = g_str_Parame & "               ELSE "
      g_str_Parame = g_str_Parame & "                    B.HIPMAE_MTOPRE "
      g_str_Parame = g_str_Parame & "               END AS CALCULADO, C.SOLMAE_MTOGCI, "
      g_str_Parame = g_str_Parame & "               (SELECT SUM(DT.DESDAT_IMPORT) FROM CRE_DESPRODAT DT "
      g_str_Parame = g_str_Parame & "                 WHERE DT.DESDAT_NUMOPE = A.DESCAB_NUMOPE AND DT.DESDAT_FECREG = DESCAB_FECREG AND DT.DESDAT_HORREG = DESCAB_HORREG) AS IMP_TOTDET, "
      g_str_Parame = g_str_Parame & "               (SELECT COUNT(*) FROM CRE_DESPRODAT DT "
      g_str_Parame = g_str_Parame & "                 WHERE DT.DESDAT_NUMOPE = A.DESCAB_NUMOPE AND DT.DESDAT_FECREG = DESCAB_FECREG AND DT.DESDAT_HORREG = DESCAB_HORREG) AS REG_TOTDET "
      g_str_Parame = g_str_Parame & "          FROM CRE_DESPROCAB A "
      g_str_Parame = g_str_Parame & "         INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.DESCAB_NUMOPE "
      g_str_Parame = g_str_Parame & "         INNER JOIN CRE_SOLMAE C ON C.SOLMAE_NUMERO = B.HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & "         WHERE A.DESCAB_CODEST IN ('3','6') "
      g_str_Parame = g_str_Parame & "           AND A.DESCAB_NUMOPE IN (" & r_str_NumOpe & ")) XX "
      g_str_Parame = g_str_Parame & " WHERE XX.CALCULADO - SOLMAE_MTOGCI <> XX.IMP_TOTDET OR XX.REG_TOTDET = 0 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            r_str_CadAux = r_str_CadAux & g_rst_Princi!DESCAB_NUMOPE & ", "
            g_rst_Princi.MoveNext
         Loop
         
         MsgBox "Favor de validar las siguientes operaciones: " & Mid(r_str_CadAux, 1, Len(r_str_CadAux) - 2), vbExclamation, modgen_g_str_NomPlt
               
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   r_str_CadAux = ""
   modgen_g_str_Mail_Mensaj = ""
   For r_int_NumFil = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_NumFil, 20) = "X" Then
          g_str_Parame = "usp_Actualiza_cre_desprocab("
          g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad.TextMatrix(r_int_NumFil, 11)) & "', "
          g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad.TextMatrix(r_int_NumFil, 17)) & "', "
          g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad.TextMatrix(r_int_NumFil, 18)) & "', "
          
          If p_Estado = 1 Then
             'APROBADO (DESCAB_CODAREA, DESCAB_CODEST) (DESCAB_CODEST)
             If CStr(Trim(grd_Listad.TextMatrix(r_int_NumFil, 12))) = "2" Then
                g_str_Parame = g_str_Parame & "'3', "
                g_str_Parame = g_str_Parame & "'4', "
             Else
                g_str_Parame = g_str_Parame & "'5', "
                g_str_Parame = g_str_Parame & "'8', "
             End If
          Else
             'RECHAZO
             If CStr(Trim(grd_Listad.TextMatrix(r_int_NumFil, 12))) = "2" Then
                g_str_Parame = g_str_Parame & "'2', "
                g_str_Parame = g_str_Parame & "'5', "
             Else
                g_str_Parame = g_str_Parame & "'4', "
                g_str_Parame = g_str_Parame & "'9', "
             End If
          End If
          g_str_Parame = g_str_Parame & "'', " 'fecha solicitud notaria
          g_str_Parame = g_str_Parame & "'', " 'comentario legal 1
          g_str_Parame = g_str_Parame & "'', " 'fecha entrega notaria
          g_str_Parame = g_str_Parame & "'', " 'comentario legal 2
          g_str_Parame = g_str_Parame & "'', " 'fecha de recepcion 2
          g_str_Parame = g_str_Parame & "'', " 'COMENTARIO.TEXT
          g_str_Parame = g_str_Parame & "'', "
          g_str_Parame = g_str_Parame & "'', "
          
          g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLE2
          g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
          g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
          g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
          g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
          
          If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
          End If
      
          modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & gf_Formato_NumSol(Trim(grd_Listad.TextMatrix(r_int_NumFil, 8))) & Chr(13)
          modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE OPERACION : " & gf_Formato_NumOpe(Trim(grd_Listad.TextMatrix(r_int_NumFil, 11))) & Chr(13)
          modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & Trim(grd_Listad.TextMatrix(r_int_NumFil, 9)) & "-" & Trim(grd_Listad.TextMatrix(r_int_NumFil, 10)) & Chr(13)
          modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & Trim(grd_Listad.TextMatrix(r_int_NumFil, 2)) & Chr(13)
          modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
       End If
   Next
         
   If p_Estado = 1 Then
      modgen_g_str_Mail_Asunto = "PAGO PROMOTOR - AREA OPERACIONES - APROBACION (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   Else
      modgen_g_str_Mail_Asunto = "PAGO PROMOTOR - AREA OPERACIONES - RECHAZO (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   End If
   
   Call fs_Envia_Correo_Prom(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, "", "", False, True, False, False, True, True)

   MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
   frm_RegDes_01.fs_Buscar_Creditos
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe datos.", vbExclamation, modgen_g_str_NomPlt
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
   moddat_g_str_NumOpe = ""
   moddat_g_str_NumSol = ""
   moddat_g_int_TipDoc = 0
   moddat_g_str_NumDoc = ""
   moddat_g_str_NomPrd = ""
   moddat_g_str_CodIte = ""
   moddat_g_int_CodIns = 0
   moddat_g_str_CodPrd = ""
   moddat_g_str_CodSub = ""
   moddat_g_int_TipMon = 0
   moddat_g_str_NomCli = ""
   moddat_g_str_FecRec = ""
   moddat_g_str_FecHip = ""
   moddat_g_str_Situac = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   Call Estado_Ctrl
   
   cmb_Estado.ListIndex = -1
   chk_Estado.Value = 1
   cmb_Estado.Enabled = False
   
   Call gs_SetFocus(cmb_Estado)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SegSol_Click()
   Dim r_str_CodIns As String
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 13
   r_str_CodIns = Trim(grd_Listad.Text)
   If r_str_CodIns <> "3" And r_str_CodIns <> "6" Then
      MsgBox "No se puede editar este registro.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   '000003 = aprobado legal, 000006 = aprobado tesoreria
   If (r_str_CodIns = "3" Or r_str_CodIns = "6") Then
      moddat_g_int_TipRep = 1
      Else
      moddat_g_int_TipRep = 2
   End If

   'numero de operacion
   grd_Listad.Col = 11
   moddat_g_str_NumOpe = Trim(grd_Listad.Text)
   'numero de solicitud
   grd_Listad.Col = 8
   moddat_g_str_NumSol = Trim(grd_Listad.Text)
   'tipo de documento
   grd_Listad.Col = 9
   moddat_g_int_TipDoc = Trim(grd_Listad.Text)
   'numero de documento
   grd_Listad.Col = 10
   moddat_g_str_NumDoc = Trim(grd_Listad.Text)
   'Nombre producto
   grd_Listad.Col = 0
   moddat_g_str_NomPrd = Trim(grd_Listad.Text)
   'Codigo de Item
   grd_Listad.Col = 13
   moddat_g_str_CodIte = Trim(grd_Listad.Text)
   'codigo de area
   grd_Listad.Col = 12
   moddat_g_int_CodIns = Trim(grd_Listad.Text)
   'HIPMAE_CODPRD
   grd_Listad.Col = 14
   moddat_g_str_CodPrd = Trim(grd_Listad.Text)
   'HIPMAE_CODSUB
   grd_Listad.Col = 15
   moddat_g_str_CodSub = Trim(grd_Listad.Text)
   'HIPMAE_MONEDA
   grd_Listad.Col = 16
   moddat_g_int_TipMon = Trim(grd_Listad.Text)
   'Nombre del Cliente
   grd_Listad.Col = 2
   moddat_g_str_NomCli = Trim(grd_Listad.Text)
   'Fecha Registro
   grd_Listad.Col = 17
   moddat_g_str_FecRec = Trim(grd_Listad.Text)
   'Hora Registro
   grd_Listad.Col = 18
   moddat_g_str_FecHip = Trim(grd_Listad.Text)
   'Estado Actual
   grd_Listad.Col = 4
   moddat_g_str_Situac = Trim(Trim(grd_Listad.Text))
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_RegDes_02.Show 1
   
   Call Estado_Ctrl
End Sub

Private Sub cmd_Seguimiento_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   '1 = Guardar = 1, Eliminar = 2, (Aprobar = 3, Rechazar = 4)
   moddat_g_int_TipRep = 2

   'numero de operacion
   grd_Listad.Col = 11
   moddat_g_str_NumOpe = Trim(grd_Listad.Text)
   'numero de solicitud
   grd_Listad.Col = 8
   moddat_g_str_NumSol = Trim(grd_Listad.Text)
   'tipo de documento
   grd_Listad.Col = 9
   moddat_g_int_TipDoc = Trim(grd_Listad.Text)
   'numero de documento
   grd_Listad.Col = 10
   moddat_g_str_NumDoc = Trim(grd_Listad.Text)
   'Nombre producto
   grd_Listad.Col = 0
   moddat_g_str_NomPrd = Trim(grd_Listad.Text)
   'Codigo de Item
   grd_Listad.Col = 13
   moddat_g_str_CodIte = Trim(grd_Listad.Text)
   '-------------------------------------------------------
   'HIPMAE_CODPRD
   grd_Listad.Col = 14
   moddat_g_str_CodPrd = Trim(grd_Listad.Text)
   'HIPMAE_CODSUB
   grd_Listad.Col = 15
   moddat_g_str_CodSub = Trim(grd_Listad.Text)
   'HIPMAE_MONEDA
   grd_Listad.Col = 16
   moddat_g_int_TipMon = Trim(grd_Listad.Text)
   'Nombre del Cliente
   grd_Listad.Col = 2
   moddat_g_str_NomCli = Trim(grd_Listad.Text)
   
   'Fecha Registro
   grd_Listad.Col = 17
   moddat_g_str_FecRec = Trim(grd_Listad.Text)
   'Hora Registro
   grd_Listad.Col = 18
   moddat_g_str_FecHip = Trim(grd_Listad.Text)
   'Estado Actual
   grd_Listad.Col = 4
   moddat_g_str_Situac = Trim(Trim(grd_Listad.Text))
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_RegDes_03.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia

   Call cmd_Limpia_Click
   chk_Estado.Value = 1
   Call chk_Estado_Click
   
   Call gs_SetFocus(cmd_Buscar)
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub Estado_Ctrl()
   If grd_Listad.Rows = 0 Then
      cmd_Seguimiento.Enabled = False
      cmd_ExpExc.Enabled = False
      cmd_SegSol.Enabled = False
      cmd_Aprobar.Enabled = False
      cmd_Rechazar.Enabled = False
      cmb_Estado.Enabled = True
      chk_Estado.Enabled = True
   Else
      cmd_Seguimiento.Enabled = True
      cmd_ExpExc.Enabled = True
      cmd_SegSol.Enabled = True
      cmd_Aprobar.Enabled = True
      cmd_Rechazar.Enabled = True
      cmb_Estado.Enabled = False
      chk_Estado.Enabled = False
   End If
End Sub

Private Sub fs_Inicia()
   cmb_Estado.Clear
   
   g_str_Parame = " SELECT to_number(PARDES_CODITE)||' - '||trim(PARDES_DESCRI) as glosa, PARDES_CODITE as codigo "
   g_str_Parame = g_str_Parame & " FROM MNT_PARDES WHERE PARDES_CODGRP = '374' "
   g_str_Parame = g_str_Parame & " and PARDES_CODITE <> '000000' AND PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " and PARDES_CODITE in (2,4) "
   g_str_Parame = g_str_Parame & " ORDER BY PARDES_CODITE ASC "
   
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
      cmb_Estado.AddItem Trim$(g_rst_Genera!glosa)
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   '---------------------------------------------
   grd_Listad.ColWidth(0) = 3420 'PRODUCTO
   grd_Listad.ColWidth(1) = 1320 'OPERACION
   grd_Listad.ColWidth(2) = 3500 'CLIENTE
   grd_Listad.ColWidth(3) = 1110 'FECHA_REGISTRO
   grd_Listad.ColWidth(4) = 2680 'ESTADO
   grd_Listad.ColWidth(5) = 1440 'CONSEJEROS
   grd_Listad.ColWidth(6) = 0 'OPERACION
   grd_Listad.ColWidth(7) = 0 'FECHA_REGISTRO
   grd_Listad.ColWidth(8) = 0 'HIPMAE_NUMSOL
   grd_Listad.ColWidth(9) = 0 'INSTANCIA
   grd_Listad.ColWidth(10) = 0 'HIPMAE_TDOCLI
   grd_Listad.ColWidth(11) = 0 'OPERACIONES
   grd_Listad.ColWidth(12) = 0 'DESCAB_CODAREA
   grd_Listad.ColWidth(13) = 0 'DESCAB_CODEST
   grd_Listad.ColWidth(14) = 0 'HIPMAE_CODPRD
   grd_Listad.ColWidth(15) = 0 'HIPMAE_CODSUB
   grd_Listad.ColWidth(16) = 0 'HIPMAE_MONEDA
   grd_Listad.ColWidth(17) = 0 '
   grd_Listad.ColWidth(18) = 0 '
   grd_Listad.ColWidth(19) = 0 '
   grd_Listad.ColWidth(20) = 1050 'SELECCION
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(20) = flexAlignCenterCenter
End Sub

Private Sub cmb_Estado_Click()
   Call Estado_Ctrl
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Estado_Click
   End If
End Sub

Public Sub cmd_Buscar_Click()
   If chk_Estado.Value = 0 Then
      If cmb_Estado.ListIndex = -1 Then
         MsgBox "Debe seleccionar una instancia.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Estado)
         Exit Sub
      End If
   End If
      
   Screen.MousePointer = 11
   Call fs_Buscar_Creditos
   
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
'   Call cmd_SegSol_Click
   
   If grd_Listad.TextMatrix(grd_Listad.Row, 20) = "X" Then
      grd_Listad.TextMatrix(grd_Listad.Row, 20) = ""
   Else
      grd_Listad.TextMatrix(grd_Listad.Row, 20) = "X"
   End If
End Sub

Private Sub grd_Listad_SelChange()
   Dim r_str_CodIns As String
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   grd_Listad.Col = 12
   r_str_CodIns = Trim(grd_Listad.Text)
   
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
End Sub

Public Sub fs_Buscar_Creditos()
Dim r_int_FlgIn1     As Integer
Dim r_int_FlgIn2     As Integer

   g_str_Parame = "  "
   g_str_Parame = g_str_Parame & "SELECT TRIM(E.PRODUC_DESCRI) AS PRODUCTO, "
   g_str_Parame = g_str_Parame & "       TRIM(A.DESCAB_NUMOPE) AS OPERACION, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
   g_str_Parame = g_str_Parame & "       A.DESCAB_FECREG AS FECHA_REGISTRO, A.DESCAB_FECREG, A.DESCAB_HORREG, "
   g_str_Parame = g_str_Parame & "       TRIM(D.PARDES_DESCRI) AS INSTANCIA, descab_codarea, descab_codest, DESCAB_CODEST, "
   g_str_Parame = g_str_Parame & "       TRIM(F.PARDES_DESCRI) AS ESTADO, B.HIPMAE_CODPRD, B.HIPMAE_CODSUB, B.HIPMAE_MONEDA, "
   g_str_Parame = g_str_Parame & "       B.hipmae_numsol, B.HIPMAE_TDOCLI, B.HIPMAE_NDOCLI, "
   g_str_Parame = g_str_Parame & "       trim(B.HIPMAE_CONHIP) AS CONSEJERO "
   g_str_Parame = g_str_Parame & "  FROM CRE_DESPROCAB A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.DESCAB_NUMOPE "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND C.DATGEN_NUMDOC = B.HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 374 AND D.PARDES_CODITE = A.DESCAB_CODAREA "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC E ON E.PRODUC_CODIGO = B.HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = 375 AND F.PARDES_CODITE = A.DESCAB_CODEST "
   
   If (cmb_Estado.ListIndex = -1 And chk_Estado.Value = False) Then
       Exit Sub
   End If
   
   If (chk_Estado.Value = False) Then
      g_str_Parame = g_str_Parame & " WHERE A.DESCAB_CODAREA = " & Trim(Mid(cmb_Estado.Text, 1, InStr(1, cmb_Estado.Text, "-") - 1)) & ""
      g_str_Parame = g_str_Parame & "   AND A.DESCAB_CODEST in ('3','6')  "
   Else
      g_str_Parame = g_str_Parame & " WHERE A.DESCAB_CODEST in ('3','6')  "
   End If
   
   g_str_Parame = g_str_Parame & " ORDER BY NOMBRE_CLIENTE "
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Operaciones para esa selección.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
    
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!PRODUCTO)
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_Formato_NumOpe(Trim(g_rst_Princi!OPERACION & ""))
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!NOMBRE_CLIENTE)
      
      grd_Listad.Col = 3
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!FECHA_REGISTRO))
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!INSTANCIA)
      
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!CONSEJERO)
      
      grd_Listad.Col = 6
      grd_Listad.Text = CStr(g_rst_Princi!OPERACION)
      
      grd_Listad.Col = 7
      grd_Listad.Text = CStr(g_rst_Princi!FECHA_REGISTRO)
      '-----------------
      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!hipmae_numsol)
      
      grd_Listad.Col = 9
      grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI)
      
      grd_Listad.Col = 10
      grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_NDOCLI)
      
      grd_Listad.Col = 11
      grd_Listad.Text = Trim(g_rst_Princi!OPERACION)
      
      grd_Listad.Col = 12
      grd_Listad.Text = Trim(g_rst_Princi!DESCAB_CODAREA)
      
      grd_Listad.Col = 13
      grd_Listad.Text = Trim(g_rst_Princi!DESCAB_CODEST)
      '--------------------------------------------------------
      grd_Listad.Col = 14
      grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_CODPRD)
      grd_Listad.Col = 15
      grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_CODSUB)
      grd_Listad.Col = 16
      grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_MONEDA)
      
      grd_Listad.Col = 17
      grd_Listad.Text = Trim(g_rst_Princi!DESCAB_FECREG)
      grd_Listad.Col = 18
      grd_Listad.Text = Trim(g_rst_Princi!DESCAB_HORREG)
            
      g_rst_Princi.MoveNext
   Loop
   
   'Ordenando por Nombre de Clientes
   pnl_Tit_NomCli.Tag = "A"
   Call gs_SorteaGrid(grd_Listad, 3, "C")
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Call gs_UbiIniGrid(grd_Listad)
   
   Call Estado_Ctrl
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Cells(2, 1) = "SEGUIMIENTO DE OPERACIONES A DESEMBOLSO AL PROMOTOR"
      .Range(.Cells(2, 1), .Cells(2, 6)).Merge
      .Range(.Cells(2, 1), .Cells(2, 6)).Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 6)).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = 4
      .Cells(r_int_NroFil, 1) = "PRODUCTO":              .Columns("A").ColumnWidth = 40
      .Cells(r_int_NroFil, 2) = "NRO OPERACION":         .Columns("B").ColumnWidth = 16
      .Cells(r_int_NroFil, 3) = "APELLIDOS Y NOMBRES":   .Columns("C").ColumnWidth = 40
      .Cells(r_int_NroFil, 4) = "FECHA REGISTRO":        .Columns("D").ColumnWidth = 15
      .Cells(r_int_NroFil, 5) = "INSTANCIA":             .Columns("E").ColumnWidth = 25
      .Cells(r_int_NroFil, 6) = "CONSEJERO":             .Columns("F").ColumnWidth = 20
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 6)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 6)).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = r_int_NroFil + 1
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 4) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         r_int_NroFil = r_int_NroFil + 1
      Next
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_old()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Cells(2, 1) = "SEGUIMIENTO DE OPERACIONES A DESEMBOLSO AL PROMOTOR"
      .Range(.Cells(2, 1), .Cells(2, 6)).Merge
      .Range(.Cells(2, 1), .Cells(2, 6)).Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 6)).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = 4
      .Cells(r_int_NroFil, 1) = "PRODUCTO":              .Columns("A").ColumnWidth = 40
      .Cells(r_int_NroFil, 2) = "NRO OPERACION":         .Columns("B").ColumnWidth = 16
      .Cells(r_int_NroFil, 3) = "APELLIDOS Y NOMBRES":   .Columns("C").ColumnWidth = 40
      .Cells(r_int_NroFil, 4) = "FECHA REGISTRO":        .Columns("D").ColumnWidth = 15
      .Cells(r_int_NroFil, 5) = "INSTANCIA ACTUAL":      .Columns("E").ColumnWidth = 25
      .Cells(r_int_NroFil, 6) = "CONSEJERO":             .Columns("F").ColumnWidth = 20
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 6)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 6)).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = r_int_NroFil + 1
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         r_int_NroFil = r_int_NroFil + 1
      Next
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub pnl_Tit_ConHip_Click()
   If Len(Trim(pnl_Tit_ConHip.Tag)) = 0 Or pnl_Tit_ConHip.Tag = "D" Then
      pnl_Tit_ConHip.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_ConHip.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecReg_Click()
   If Len(Trim(pnl_Tit_FecReg.Tag)) = 0 Or pnl_Tit_FecReg.Tag = "D" Then
      pnl_Tit_FecReg.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "N")
   Else
      pnl_Tit_FecReg.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "N-")
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

Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If

End Sub

Private Sub pnl_Tit_SitAct_Click()
   If Len(Trim(pnl_Tit_SitAct.Tag)) = 0 Or pnl_Tit_SitAct.Tag = "D" Then
      pnl_Tit_SitAct.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Tit_SitAct.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub
