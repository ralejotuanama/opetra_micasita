VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Desemb_19 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   4425
   ClientTop       =   2475
   ClientWidth     =   8760
   Icon            =   "OpeTra_frm_153.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5115
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8790
      _Version        =   65536
      _ExtentX        =   15505
      _ExtentY        =   9022
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
      Begin Threed.SSPanel SSPanel11 
         Height          =   2805
         Left            =   30
         TabIndex        =   5
         Top             =   2250
         Width           =   8685
         _Version        =   65536
         _ExtentX        =   15319
         _ExtentY        =   4948
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
         Begin VB.DirListBox dir_LisCar 
            Height          =   2340
            Left            =   1680
            TabIndex        =   1
            Top             =   390
            Width           =   6945
         End
         Begin VB.DriveListBox drv_LisUni 
            Height          =   315
            Left            =   1680
            TabIndex        =   0
            Top             =   60
            Width           =   6945
         End
         Begin VB.Label Label2 
            Caption         =   "Carpeta a guardar archivos:"
            Height          =   615
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   8685
         _Version        =   65536
         _ExtentX        =   15319
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
            Left            =   660
            TabIndex        =   8
            Top             =   30
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   315
            Left            =   660
            TabIndex        =   9
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Exportación de Cronogramas a Mivivienda"
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
            Picture         =   "OpeTra_frm_153.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   8685
         _Version        =   65536
         _ExtentX        =   15319
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
            Left            =   8070
            Picture         =   "OpeTra_frm_153.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Export 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_153.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Generar Cronogramas para Mivivienda"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   11
         Top             =   1440
         Width           =   8685
         _Version        =   65536
         _ExtentX        =   15319
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1680
            TabIndex        =   12
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1680
            TabIndex        =   13
            Top             =   390
            Width           =   6945
            _Version        =   65536
            _ExtentX        =   12250
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
         Begin VB.Label Label3 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   420
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   90
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Desemb_19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Export_Click()
   If MsgBox("¿Está seguro de generar los archivos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Genera_Archivos(dir_LisCar.Path)
   Screen.MousePointer = 0
   
   MsgBox "Archivo Generado correctamente.", vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub drv_LisUni_Change()
   dir_LisCar.Path = drv_LisUni.Drive
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Genera_Archivos(ByVal p_RutFil As String)
   Dim r_int_NumFil     As Integer
   Dim r_str_NomCab     As String
   Dim r_str_NomDet     As String
   Dim r_dbl_PorNCo     As Double
   Dim r_dbl_PorCon     As Double
   Dim r_dbl_TotPre     As Double
   Dim r_dbl_TotCuo     As Double
   Dim r_str_OpeMVi     As String
   Dim r_int_PerGra     As Integer
   Dim r_dbl_SalCap     As Double
   Dim r_dbl_MtoPre     As Double
   Dim r_dbl_MtoGra     As Double
   

   r_str_NomCab = p_RutFil & "\" & "C" & Format(date, "yymmdd") & ".064"
   r_str_NomDet = p_RutFil & "\" & "D" & Format(date, "yymmdd") & ".064"

   r_int_NumFil = FreeFile
   Open r_str_NomCab For Output As r_int_NumFil

   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_int_PerGra = g_rst_Princi!HIPMAE_PERGRA
      r_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOMVI
      r_dbl_MtoGra = g_rst_Princi!HIPMAE_IMPCON + g_rst_Princi!HIPMAE_IMPNCO
      
      r_dbl_PorNCo = g_rst_Princi!HIPMAE_IMPNCO / (g_rst_Princi!HIPMAE_MTOMVI + g_rst_Princi!HIPMAE_INTCAP) * 100
      r_dbl_PorNCo = CDbl(Format(r_dbl_PorNCo, "######0.0000"))
      
      r_dbl_PorCon = g_rst_Princi!HIPMAE_IMPCON / (g_rst_Princi!HIPMAE_MTOMVI + g_rst_Princi!HIPMAE_INTCAP) * 100
      r_dbl_PorCon = CDbl(Format(r_dbl_PorCon, "######0.0000"))
      
      r_dbl_TotPre = g_rst_Princi!HIPMAE_IMPCON + g_rst_Princi!HIPMAE_IMPNCO
      
      r_str_OpeMVi = Trim(g_rst_Princi!HIPMAE_OPEMVI)
      
      Print #r_int_NumFil, "1" & " " & _
                           Format(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))), "yyyymmdd") & " " & _
                           Space(16 - Len(Trim(g_rst_Princi!HIPMAE_OPEMVI))) & Trim(g_rst_Princi!HIPMAE_OPEMVI) & " " & _
                           Space(20 - Len(Trim(g_rst_Princi!HIPMAE_NUMOPE))) & Trim(g_rst_Princi!HIPMAE_NUMOPE) & " " & _
                           CStr(g_rst_Princi!HIPMAE_TDOCLI) & " " & _
                           Space(12 - Len(Trim(g_rst_Princi!HIPMAE_NDOCLI))) & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " " & _
                           Mid(moddat_gf_Consulta_ParDes("237", CStr(g_rst_Princi!HIPMAE_MONEDA)) & Space(1), 1, 1) & " " & _
                           Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_MTOMVI, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_MTOMVI, "########0.00"), 2) & " " & _
                           Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_IMPNCO, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_IMPNCO, "########0.00"), 2) & " " & _
                           Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_IMPCON, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_IMPCON, "########0.00"), 2) & " " & _
                           Space(12 - Len(gf_ComaDecimal(CStr(r_dbl_PorNCo), 4))) & gf_ComaDecimal(CStr(r_dbl_PorNCo), 4) & " " & _
                           Space(12 - Len(gf_ComaDecimal(CStr(r_dbl_PorCon), 4))) & gf_ComaDecimal(CStr(r_dbl_PorCon), 4) & " " & _
                           Format(g_rst_Princi!HIPMAE_NUMCUO + g_rst_Princi!HIPMAE_PERGRA, "000") & " " & _
                           Space(3 - Len(CStr(g_rst_Princi!HIPMAE_PERGRA))) & CStr(g_rst_Princi!HIPMAE_PERGRA) & " " & _
                           Space(12 - Len(gf_ComaDecimal(CStr(r_dbl_TotPre), 2))) & gf_ComaDecimal(CStr(r_dbl_TotPre), 2)

   End If
   
   'Cerrando Archivo Cabecera
   Close #r_int_NumFil

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'Creando Detalle
   r_int_NumFil = FreeFile
   Open r_str_NomDet For Output As r_int_NumFil

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 3 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_dbl_SalCap = g_rst_Princi!HIPCUO_SALCAP
         
         If r_int_PerGra > 0 Then
            If g_rst_Princi!HIPCUO_NUMCUO < r_int_PerGra Then
               r_dbl_SalCap = r_dbl_MtoPre
            ElseIf g_rst_Princi!HIPCUO_NUMCUO = r_int_PerGra Then
               r_dbl_SalCap = r_dbl_MtoGra
            End If
         End If
         
         r_dbl_TotCuo = g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_DESORG + g_rst_Princi!HIPCUO_VIVORG + g_rst_Princi!HIPCUO_OTRORG
      
         Print #r_int_NumFil, Space(16 - Len(Trim(r_str_OpeMVi))) & Trim(r_str_OpeMVi) & " " & _
                              Space(3 - Len(CStr(g_rst_Princi!HIPCUO_NUMCUO))) & CStr(g_rst_Princi!HIPCUO_NUMCUO) & " " & _
                              Format(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))), "yyyymmdd") & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_CAPITA, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_CAPITA, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_INTERE, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_INTERE, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_DESORG, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_DESORG, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_VIVORG, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_VIVORG, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_OTRORG, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_OTRORG, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(r_dbl_TotCuo, "########0.00"), 2))) & gf_ComaDecimal(Format(r_dbl_TotCuo, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(r_dbl_SalCap, "########0.00"), 2))) & gf_ComaDecimal(Format(r_dbl_SalCap, "########0.00"), 2)

         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Cerrando Archivo Detalle
   Close #r_int_NumFil
End Sub




