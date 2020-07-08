VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Pro_CbzCof 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10365
   Icon            =   "OpeTra_frm_349.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   5100
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10365
      _Version        =   65536
      _ExtentX        =   18283
      _ExtentY        =   8996
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
         Height          =   825
         Left            =   60
         TabIndex        =   8
         Top             =   1500
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
         _ExtentY        =   1455
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2685
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   420
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
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   315
            Left            =   150
            TabIndex        =   10
            Top             =   420
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   150
            TabIndex        =   9
            Top             =   90
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
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
            TabIndex        =   12
            Top             =   90
            Width           =   7935
            _Version        =   65536
            _ExtentX        =   13996
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Procesos - Carga de Archivo Cobranza COFIDE"
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
            Picture         =   "OpeTra_frm_349.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   60
         TabIndex        =   13
         Top             =   810
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
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
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   60
            Picture         =   "OpeTra_frm_349.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Cargar Saldos COFIDE"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9630
            Picture         =   "OpeTra_frm_349.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2175
         Left            =   60
         TabIndex        =   14
         Top             =   2370
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
         _ExtentY        =   3836
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
         Begin VB.FileListBox fil_LisArc 
            Height          =   2040
            Left            =   1560
            TabIndex        =   4
            Top             =   90
            Width           =   4425
         End
         Begin VB.DriveListBox drv_LisUni 
            Height          =   315
            Left            =   6060
            TabIndex        =   2
            Top             =   90
            Width           =   4095
         End
         Begin VB.DirListBox dir_LisCar 
            Height          =   1665
            Left            =   6060
            TabIndex        =   3
            Top             =   420
            Width           =   4095
         End
         Begin VB.Label Label1 
            Caption         =   "Archivo a cargar:"
            Height          =   315
            Left            =   150
            TabIndex        =   15
            Top             =   90
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   60
         TabIndex        =   16
         Top             =   4590
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
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
         Begin VB.Label lbl_NomPro 
            Caption         =   "Proceso carga información Cobranza COFIDE"
            Height          =   255
            Left            =   30
            TabIndex        =   17
            Top             =   120
            Width           =   5505
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_CbzCof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Proces_Click()
Dim r_lng_Contad           As Long
Dim modprc_g_str_CadEje    As String

   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   If Len(Trim(fil_LisArc.FileName & "")) = 0 Then
      MsgBox "Debe seleccionar el Archivo a cargar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      Exit Sub
   End If
         
   If MsgBox("¿Está seguro de ejecutar el proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   cmd_Proces.Enabled = False
   
   g_str_Parame = "SELECT NVL(COUNT(*),0) AS TOTREG FROM CRE_ARCMEN WHERE ARCMEN_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND ARCMEN_PERANO = " & CStr(ipp_PerAno.Text)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   r_lng_Contad = g_rst_Princi!TOTREG
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If r_lng_Contad > 0 Then
      If MsgBox("La información de cobranza COFIDE para este Período ya ha sido cargada. ¿Desea volver a cargar la información?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
         lbl_NomPro.Caption = "Eliminando información Cobranza COFIDE...": DoEvents
         
         modprc_g_str_CadEje = "DELETE FROM CRE_ARCMEN WHERE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "ARCMEN_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "ARCMEN_PERANO = " & CStr(ipp_PerAno.Text)
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, g_rst_GenAux, 2) Then
            Exit Sub
         End If
         
         lbl_NomPro.Caption = "Proceso carga información Cobranza COFIDE...": DoEvents
         Call fs_CargaCOFIDE(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text))
      End If
   Else
      lbl_NomPro.Caption = "Proceso carga información Cobranza COFIDE...": DoEvents
      Call fs_CargaCOFIDE(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text))
   End If
   
   cmd_Proces.Enabled = True
   Screen.MousePointer = 0
   MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub fs_CargaCOFIDE(ByVal p_ArcCOFIDE, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
Dim r_obj_Excel     As Excel.Application
Dim r_int_FilExc    As Integer
Dim r_int_FilTot    As Long
Dim r_int_IDCIPR    As Integer
Dim r_str_NOMPRO    As String
Dim r_str_CodCof    As String
Dim r_str_NomCli    As String
Dim r_str_NUMCTR    As String
Dim r_str_NUMALT    As String
Dim r_str_TIPMON    As String
Dim r_dbl_EXPINI    As Double
Dim r_dbl_Princi    As Double
Dim r_dbl_IMPINT    As Double
Dim r_dbl_IMTASA    As Double
Dim r_dbl_COMSIN    As Double
Dim r_dbl_ImpTot    As Double
Dim r_dbl_EXPFIN    As Double
Dim r_str_BUENPG    As String
Dim r_str_MALPAG    As String
       
    'Abriendo Archivo COFIDE
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=p_ArcCOFIDE
   r_int_FilExc = 1
   
   With r_obj_Excel.Sheets(1)
        r_int_FilTot = CStr(.Cells(.Rows.Count, 1).End(xlUp).Row)
        Do While r_int_FilTot <> r_int_FilExc
                 
           If (Len(Trim(Cells(r_int_FilExc, 1).Value)) >= 8 And Len(Trim(Cells(r_int_FilExc, 3).Value)) >= 3 And _
               Len(Trim(Cells(r_int_FilExc, 5).Value)) >= 8) Then
               If (IsNumeric(Trim(Cells(r_int_FilExc, 1).Value)) = True And IsNumeric(Trim(Cells(r_int_FilExc, 3).Value)) = True And _
                   IsNumeric(Trim(Cells(r_int_FilExc, 5).Value)) = True) Then
                   
                   r_int_IDCIPR = 0:  r_str_NOMPRO = "": r_str_CodCof = ""
                   r_str_NomCli = "": r_str_NUMCTR = "": r_str_NUMALT = "": r_str_TIPMON = ""
                   r_dbl_EXPINI = 0:  r_dbl_Princi = 0:  r_dbl_IMPINT = 0
                   r_dbl_IMTASA = 0:  r_dbl_COMSIN = 0:  r_dbl_ImpTot = 0
                   r_dbl_EXPFIN = 0:  r_str_BUENPG = 0:  r_str_MALPAG = 0
                     
                   r_int_IDCIPR = Trim(Cells(r_int_FilExc, 3).Value)
                   r_str_NOMPRO = Trim(Cells(r_int_FilExc, 4).Value)
                   r_str_CodCof = Trim(Cells(r_int_FilExc, 5).Value)
                   r_str_NomCli = Trim(Cells(r_int_FilExc, 6).Value)
                   r_str_NUMCTR = Trim(Cells(r_int_FilExc, 7).Value)
                   r_str_NUMALT = Trim(Cells(r_int_FilExc, 8).Value)
                   r_str_TIPMON = Trim(Cells(r_int_FilExc, 9).Value)
                   r_dbl_EXPINI = Trim(Cells(r_int_FilExc, 10).Value)
                   r_dbl_Princi = Trim(Cells(r_int_FilExc, 11).Value)
                   r_dbl_IMPINT = Trim(Cells(r_int_FilExc, 12).Value)
                   r_dbl_IMTASA = Trim(Cells(r_int_FilExc, 13).Value)
                   r_dbl_COMSIN = Trim(Cells(r_int_FilExc, 14).Value)
                   r_dbl_ImpTot = Trim(Cells(r_int_FilExc, 15).Value)
                   r_dbl_EXPFIN = Trim(Cells(r_int_FilExc, 16).Value)
                   r_str_BUENPG = Trim(Cells(r_int_FilExc, 17).Value)
                   r_str_MALPAG = Trim(Cells(r_int_FilExc, 18).Value)
               
                   g_str_Parame = ""
                   g_str_Parame = "INSERT INTO CRE_ARCMEN ("
                   g_str_Parame = g_str_Parame & "ARCMEN_PERANO, "
                   g_str_Parame = g_str_Parame & "ARCMEN_PERMES, "
                   g_str_Parame = g_str_Parame & "ARCMEN_NUMCTR, "
                   g_str_Parame = g_str_Parame & "ARCMEN_IDCIPR, "
                   g_str_Parame = g_str_Parame & "ARCMEN_NOMPRO, "
                   g_str_Parame = g_str_Parame & "ARCMEN_CODCOF, "
                   g_str_Parame = g_str_Parame & "ARCMEN_NOMCLI, "
                   g_str_Parame = g_str_Parame & "ARCMEN_NUMALT, "
                   g_str_Parame = g_str_Parame & "ARCMEN_TIPMON, "
                   g_str_Parame = g_str_Parame & "ARCMEN_EXPINI, "
                   g_str_Parame = g_str_Parame & "ARCMEN_PRINCI, "
                   g_str_Parame = g_str_Parame & "ARCMEN_IMPINT, "
                   g_str_Parame = g_str_Parame & "ARCMEN_IMTASA, "
                   g_str_Parame = g_str_Parame & "ARCMEN_COMSIN, "
                   g_str_Parame = g_str_Parame & "ARCMEN_IMPTOT, "
                   g_str_Parame = g_str_Parame & "ARCMEN_EXPFIN, "
                   g_str_Parame = g_str_Parame & "ARCMEN_BUENPG, "
                   g_str_Parame = g_str_Parame & "ARCMEN_MALPAG, "
                   g_str_Parame = g_str_Parame & "SEGUSUCRE, "
                   g_str_Parame = g_str_Parame & "SEGFECCRE, "
                   g_str_Parame = g_str_Parame & "SEGHORCRE, "
                   g_str_Parame = g_str_Parame & "SEGPLTCRE, "
                   g_str_Parame = g_str_Parame & "SEGTERCRE, "
                   g_str_Parame = g_str_Parame & "SEGSUCCRE) "
                   g_str_Parame = g_str_Parame & "VALUES ( "
                   g_str_Parame = g_str_Parame & p_PerAno & " , "
                   g_str_Parame = g_str_Parame & p_PerMes & " , "
                   g_str_Parame = g_str_Parame & "'" & r_str_NUMCTR & "', "
                   g_str_Parame = g_str_Parame & "'" & r_int_IDCIPR & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_NOMPRO & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_CodCof & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_NomCli & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_NUMALT & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_TIPMON & "', "
                   g_str_Parame = g_str_Parame & r_dbl_EXPINI & ", "
                   g_str_Parame = g_str_Parame & r_dbl_Princi & ", "
                   g_str_Parame = g_str_Parame & r_dbl_IMPINT & ", "
                   g_str_Parame = g_str_Parame & r_dbl_IMTASA & ", "
                   g_str_Parame = g_str_Parame & r_dbl_COMSIN & ", "
                   g_str_Parame = g_str_Parame & r_dbl_ImpTot & ", "
                   g_str_Parame = g_str_Parame & r_dbl_EXPFIN & ", "
                   g_str_Parame = g_str_Parame & "'" & r_str_BUENPG & "', "
                   g_str_Parame = g_str_Parame & "'" & r_str_MALPAG & "', "
               
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
                   g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
                   g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & ", "
                   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "' ,"
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
               
                   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                      Exit Sub
                   End If
                   
               End If
           End If
           r_int_FilExc = r_int_FilExc + 1
        Loop
   End With
   
   r_obj_Excel.Workbooks.Close
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub dir_LisCar_Change()
   fil_LisArc.Path = dir_LisCar.Path
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Value = Year(date)
   
   drv_LisUni.Drive = "C:"
   dir_LisCar.Path = "C:\"
End Sub

Private Sub fs_Limpia()
Dim r_int_PerMes  As Integer
Dim r_int_PerAno  As Integer

   If Month(date) = 1 Then
      r_int_PerMes = 12
      r_int_PerAno = Year(date) - 1
   Else
      r_int_PerMes = Month(date) - 1
      r_int_PerAno = Year(date)
   End If

   ipp_PerAno.Text = Format(r_int_PerAno, "0000")
   fil_LisArc.Pattern = "*.xls"
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(drv_LisUni)
   End If
End Sub

Private Sub drv_LisUni_Change()
   dir_LisCar.Path = drv_LisUni.Drive
End Sub

Private Sub drv_LisUni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error Resume Next
      dir_LisCar.Path = drv_LisUni.Drive
   End If
End Sub
