VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Caj_GenCro_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   4200
   ClientLeft      =   2610
   ClientTop       =   3465
   ClientWidth     =   7905
   Icon            =   "OpeTra_frm_332.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      _Version        =   65536
      _ExtentX        =   13944
      _ExtentY        =   7646
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
         TabIndex        =   12
         Top             =   770
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
            Picture         =   "OpeTra_frm_332.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Import 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_332.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Importar el Cronograma"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1485
         Left            =   30
         TabIndex        =   13
         Top             =   2655
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   2619
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
         Begin VB.ComboBox cmb_TipPre 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   390
            Width           =   2715
         End
         Begin VB.TextBox txt_NroCuo 
            Height          =   315
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   9
            Text            =   "txt_NoCuo2"
            Top             =   1050
            Width           =   915
         End
         Begin VB.TextBox txt_NomArc 
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "txt_NomArc"
            Top             =   720
            Width           =   5835
         End
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
            TabIndex        =   8
            Top             =   720
            Width           =   315
         End
         Begin VB.ComboBox cmb_TipCro 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   60
            Width           =   2715
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo Prepago:"
            Height          =   315
            Left            =   90
            TabIndex        =   24
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label lbl_NroCuo 
            Caption         =   "Nro de cuota TC2:"
            Height          =   315
            Left            =   90
            TabIndex        =   19
            Top             =   1100
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Archivo a cargar:"
            Height          =   255
            Left            =   90
            TabIndex        =   15
            Top             =   760
            Width           =   1365
         End
         Begin VB.Label Label3 
            Caption         =   "Cronograma:"
            Height          =   315
            Left            =   90
            TabIndex        =   14
            Top             =   120
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   30
         TabIndex        =   16
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
            Left            =   660
            TabIndex        =   17
            Top             =   60
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
            Left            =   660
            TabIndex        =   18
            Top             =   360
            Width           =   4845
            _Version        =   65536
            _ExtentX        =   8546
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Carga de los Cronogramas 2,3,4 y 5"
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
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   7200
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_332.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1160
         Left            =   30
         TabIndex        =   20
         Top             =   1455
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   2028
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.21
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   7200
            Picture         =   "OpeTra_frm_332.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   540
            Width           =   585
         End
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   400
            Width           =   1390
            _Version        =   65536
            _ExtentX        =   2452
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "pnl_DocIde"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1560
            TabIndex        =   4
            Top             =   750
            Width           =   5475
            _Version        =   65536
            _ExtentX        =   9657
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "pnl_Client"
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
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   60
            Width           =   1390
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   90
            TabIndex        =   23
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Top             =   790
            Width           =   1125
         End
         Begin VB.Label Label20 
            Caption         =   "Doc. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   470
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frm_Caj_GenCro_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   

Private Sub cmb_TipCro_Click()
    txt_NroCuo.Text = ""
    Call gs_SetFocus(txt_NroCuo)
    Select Case CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex))
        Case 2: lbl_NroCuo.Caption = "Nro de cuota TC2:"
        Case 3: lbl_NroCuo.Caption = "Nro de cuota TC3:"
        Case 4: lbl_NroCuo.Caption = "Nro de cuota TC4:"
        Case 5: lbl_NroCuo.Caption = "Nro de cuota TC5:"
    End Select
End Sub
 

Private Sub cmd_BusArc_Click()
    On Error GoTo cmd_BusArc_Error
    
    dlg_Guarda.Filter = "Archivos Excel |*.xls"
    dlg_Guarda.ShowOpen
    txt_NomArc.Text = UCase(dlg_Guarda.FileName)
    Exit Sub
    
cmd_BusArc_Error:
    txt_NomArc.Text = ""
End Sub

Private Sub cmd_Buscar_Click()
    '1 ok / 2 error
    moddat_g_int_FlgGOK = 1
    If Len(Trim(msk_NumOpe.Text)) < 10 Then
        MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_con_OpeTra
        Call gs_SetFocus(msk_NumOpe)
        moddat_g_int_FlgGOK = 2
        Exit Sub
    End If
    
    g_str_Parame = "  "
    g_str_Parame = g_str_Parame & " SELECT  "
    g_str_Parame = g_str_Parame & " TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRE "
    g_str_Parame = g_str_Parame & " , DATGEN_TIPDOC  ||'-' || DATGEN_NUMDOC AS DNI"
    g_str_Parame = g_str_Parame & " FROM CLI_DATGEN "
    g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE ON (DATGEN_NUMDOC = HIPMAE_NDOCLI AND DATGEN_TIPDOC = HIPMAE_TDOCLI)"
    g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & Trim(msk_NumOpe.Text) & "'"
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
        moddat_g_int_FlgGOK = 2
        Exit Sub
    End If
    
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
        MsgBox "No existe ningún cliente registrado con ese Número. ", vbExclamation, modgen_g_con_OpeTra
        msk_NumOpe.Text = ""
        pnl_DocIde.Caption = ""
        pnl_Client.Caption = ""
        Call gs_SetFocus(msk_NumOpe)
        moddat_g_int_FlgGOK = 2
        g_rst_Princi.Close
        Set g_rst_Princi = Nothing
        Exit Sub
    End If

    pnl_DocIde.Caption = g_rst_Princi!DNI
    pnl_Client.Caption = g_rst_Princi!NOMBRE
    
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
    
End Sub

Private Sub fs_Limpiar()
    msk_NumOpe.Text = ""
    pnl_DocIde.Caption = ""
    pnl_Client.Caption = ""
    txt_NomArc.Text = ""
    txt_NroCuo.Text = ""
    
    Call gs_SetFocus(msk_NumOpe)
End Sub

Private Sub cmd_Import_Click()
    If Len(Trim(msk_NumOpe.Text)) = 0 Then
        MsgBox "Debe ingresar el número de Operación.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(msk_NumOpe)
        Exit Sub
    End If
    
    If cmb_TipCro.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de Cronograma.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_TipCro)
        Exit Sub
    End If
    
    If cmb_TipPre.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de Prepago.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmb_TipPre)
        Exit Sub
    End If
    
    If Len(Trim(txt_NomArc.Text)) = 0 Then
        MsgBox "Debe ingresar la ubicación y nombre del archivo a importar.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(txt_NomArc)
        Exit Sub
    End If
    
    If Len(Trim(txt_NroCuo.Text)) = 0 Then
        MsgBox "Debe ingresar el Nro de Cuota.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(txt_NroCuo)
        Exit Sub
    End If
        
    If MsgBox("¿Está seguro de ejecutar el proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    Me.Enabled = False
    Select Case CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex))
        Case 2: Call fs_Cronograma(txt_NomArc.Text, 2)
        Case 3: Call fs_Cronograma(txt_NomArc.Text, 3)
        Case 4: Call fs_Cronograma(txt_NomArc.Text, 4)
        Case 5: Call fs_Cronograma(txt_NomArc.Text, 5)
    End Select
    Me.Enabled = True
    Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Caption = modgen_g_str_NomPlt
    
    Call fs_Inicia
    Call fs_Limpiar
    Call gs_CentraForm(Me)
    
    Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
    'TIPOS DE CRONOGRAMAS
    Call moddat_gs_Carga_LisIte_Combo(cmb_TipCro, 1, "048")
    cmb_TipCro.RemoveItem (0)
    cmb_TipCro.ListIndex = -1
    
    'TIPOS DE PREPAGOS
    cmb_TipPre.Clear
    cmb_TipPre.AddItem "REDUCCION DE MONTO"
    cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 1
    cmb_TipPre.AddItem "REDUCCION DE PLAZO"
    cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 2
    cmb_TipPre.ListIndex = 0
End Sub

'***********************************************************
'***************** IMPORTAR CRONOGRAMA *********************
'***********************************************************
 
Private Sub fs_Cronograma(ByVal RutaCr As String, ByVal r_int_TipCro As Integer)
    Dim r_obj_Excel     As excel.Application
    Dim r_int_NroCro    As Integer
    Dim r_int_nrofil    As Integer
    Dim r_int_CuoExc    As Integer
    Dim r_int_NroCol    As Integer
    
 
    Call cmd_Buscar_Click
    If moddat_g_int_FlgGOK = 2 Then Exit Sub
    
    r_int_nrofil = 2
    r_int_NroCro = CInt(Trim(txt_NroCuo.Text))
    
    Set r_obj_Excel = New excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
 
    r_obj_Excel.Workbooks.Open FileName:=RutaCr
    
    
    If (CInt(Mid(msk_NumOpe.Text, 1, 3)) = 3 And CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex)) <> 5) Then
        r_int_NroCol = 2
    Else
        r_int_NroCol = IIf(CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex)) = 2, 2, 3)
    End If
    
    'obtener la ultima cuota del excel
    Do While Not Trim(r_obj_Excel.Cells(r_int_nrofil, r_int_NroCol).Value) = ""
        r_int_CuoExc = Trim(r_obj_Excel.Cells(r_int_nrofil, r_int_NroCol).Value)
        r_int_nrofil = r_int_nrofil + 1
    Loop
    
    r_int_nrofil = 2
    
    'VERIFICAR SI LA CUOTA ESTA O NO CANCELADA
    If fs_Verificar_CuoPag(txt_NroCuo.Text, r_int_CuoExc) Then
    
        'TIPO DE PREPAGO
        If CInt(cmb_TipPre.ItemData(cmb_TipPre.ListIndex)) = 2 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "USP_CRONOG_UPDDEL("
            g_str_Parame = g_str_Parame & "'" & Trim(msk_NumOpe.Text) & "', "
            g_str_Parame = g_str_Parame & "'" & r_int_TipCro & "')"

            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 2) Then
                MsgBox "No se pudo completar el procedimiento USP_CRONOG_UPDDEL.", vbExclamation, modgen_g_str_NomPlt
                r_obj_Excel.Quit
                Set r_obj_Excel = Nothing
                Exit Sub
            End If
        End If
       
        'FLAG 1: GRABADO, 2: NO GRABADO
        moddat_g_int_FlgGrb = 2
       
        'EJECUTAR LA ACTUALIZACION DE LOS CRONOGRAMAS
        Do While Not Trim(r_obj_Excel.Cells(r_int_nrofil, r_int_NroCol).Value) = ""
    
            If CInt(r_obj_Excel.Cells(r_int_nrofil, r_int_NroCol).Value) >= r_int_NroCro Then
                moddat_g_int_FlgGrb = 1
                g_str_Parame = ""
                g_str_Parame = g_str_Parame & "USP_CRONOG_UPDATE("
                g_str_Parame = g_str_Parame & "'" & Trim(msk_NumOpe.Text) & "', "
                g_str_Parame = g_str_Parame & CInt(r_obj_Excel.Cells(r_int_nrofil, r_int_NroCol).Value) & ", "
                
                Select Case CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex))
                    Case 2:
                        g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 5).Value) & ", "     'CAPITAL
                        g_str_Parame = g_str_Parame & Format(CDbl(r_obj_Excel.Cells(r_int_nrofil, 6).Value), "######.##") & ", "    'INTERES
                        g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 7).Value) & ", "     '
                        g_str_Parame = g_str_Parame & "0,0,0, "
                        g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 8).Value) & ", "     'SALDO CAPITAL
                    Case 3, 4, 5:
                    
                        If (CInt(Mid(msk_NumOpe.Text, 1, 3)) = 3 And CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex)) = 3) Then
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 4).Value) & ", "     'CAPITAL
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 5).Value) & ", "     'INTERES
                            g_str_Parame = g_str_Parame & "0, "                                                     'DES. ORG.
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 6).Value) & ", "     '
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 7).Value) & ", "     '
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 8).Value) & ", "     '
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 10).Value) & ", "    'SALDO CAPITAL
                        
                        ElseIf (CInt(Mid(msk_NumOpe.Text, 1, 3)) = 3 And CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex)) = 4) Then
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 4).Value) & ", "     'CAPITAL
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 5).Value) & ", "     'INTERES
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 6).Value) & ", "     'DES. ORG.
                            g_str_Parame = g_str_Parame & "0, 0, 0,"
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 8).Value) & ", "    'SALDO CAPITAL
                        
                        Else
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 7).Value) & ", "     'CAPITAL
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 8).Value) & ", "     'INTERES
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 9).Value) & ", "     'DES. ORG.
                            g_str_Parame = g_str_Parame & "0,0,0, "
                            g_str_Parame = g_str_Parame & CDbl(r_obj_Excel.Cells(r_int_nrofil, 11).Value) & ", "    'SALDO CAPITAL
                        End If
                End Select
            
                g_str_Parame = g_str_Parame & "'" & r_int_TipCro & "', "
                'Datos de Auditoria
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                          'Código Usuario
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                          'Nombre Terminal
                g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                           'Nombre Ejecutable
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                          'Código Sucursal
                g_str_Parame = g_str_Parame & "'" & r_int_CuoExc & "', "
                g_str_Parame = g_str_Parame & CInt(cmb_TipPre.ItemData(cmb_TipPre.ListIndex)) & ") "
    
                If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                    MsgBox "No se pudo completar el procedimiento USP_CRONOG_UPDATE.", vbExclamation, modgen_g_str_NomPlt
                    r_obj_Excel.Quit
                    Set r_obj_Excel = Nothing
                    Exit Sub
                End If
            End If
            
            r_int_nrofil = r_int_nrofil + 1
        Loop
         
        
        If moddat_g_int_FlgGrb = 1 Then
            Call fs_Insert_LogCro(r_int_CuoExc)
            
            'CREAR UN LOG DEL ARCHIVO DE EXCEL EN LA RUTA " \\Server_micasita\APLICACIONES\Logs "
            Call fs_LogExc(r_obj_Excel)
            
            MsgBox "Se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
        Else
             MsgBox "No se encontró la cuota.", vbExclamation, modgen_g_str_NomPlt
        End If
        
        
    End If
    
    r_obj_Excel.Quit
    Set r_obj_Excel = Nothing
 
End Sub


'TABLA LOG DE ACTUALIZACION DE CRONOGRAMAS
Private Sub fs_Insert_LogCro(ByVal r_int_CuoExc As Integer)
      
    g_str_Parame = "USP_CRONOG_LOG("
    g_str_Parame = g_str_Parame & "'" & Trim(msk_NumOpe.Text) & "', "
    g_str_Parame = g_str_Parame & CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex)) & ", "
    g_str_Parame = g_str_Parame & CInt(cmb_TipPre.ItemData(cmb_TipPre.ListIndex)) & ", "
    g_str_Parame = g_str_Parame & "'" & Trim(txt_NomArc.Text) & "', "
    
    'FLAG DE CUOTA YA CANCELADA -> 1: no cancelada; 2: cancelada
    g_str_Parame = g_str_Parame & CInt(moddat_g_int_CntErr) & ", "
                            
    'Datos de Auditoria
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "         'Código Usuario
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "         'Nombre Terminal
    g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "          'Nombre Ejecutable
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "         'Código Sucursal
    
    g_str_Parame = g_str_Parame & CInt(txt_NroCuo.Text) & ", "
    g_str_Parame = g_str_Parame & r_int_CuoExc & ") "
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
        MsgBox "No se pudo completar el procedimiento USP_CRONOG_LOG.", vbExclamation, modgen_g_str_NomPlt
        Exit Sub
    End If
End Sub

 
'LOG DEL ARCHIVO DE EXCEL EN LA RUTA
Private Sub fs_LogExc(ByVal r_obj_Excel As excel.Application)
    Dim r_obj_Workbk    As excel.Workbook
    Dim r_str_SaveAs    As String
     
    r_str_SaveAs = Trim(txt_NomArc.Text)
    Do While InStr(r_str_SaveAs, "\") > 0
        r_str_SaveAs = Right(r_str_SaveAs, Len(r_str_SaveAs) - InStr(r_str_SaveAs, "\"))
    Loop
    
    r_str_SaveAs = "\\Server_micasita\APLICACIONES\Logs\" & Format(Now, "yyyymmdd") & "-" & Format(Now, "hhmmss") & "_" & r_str_SaveAs
  
    Set r_obj_Workbk = r_obj_Excel.ActiveWorkbook
    r_obj_Excel.DisplayAlerts = False
 
    r_obj_Workbk.SaveAs FileName:= _
        r_str_SaveAs, FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False
 
    r_obj_Workbk.Close SaveChanges:=False
End Sub


'VERIFICAR SI LA CUOTA ESTA O NO CANCELADA
Function fs_Verificar_CuoPag(ByVal r_int_NroCuo As Integer, ByVal r_int_CuoExc As Integer) As Boolean
    fs_Verificar_CuoPag = True
    moddat_g_int_CntErr = 1
    
    If CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex)) = 5 Then Exit Function
    
    g_str_Parame = "  "
    g_str_Parame = g_str_Parame & "SELECT HIPCUO_SITUAC, HIPCUO_NUMCUO "
    g_str_Parame = g_str_Parame & "FROM CRE_HIPCUO "
    g_str_Parame = g_str_Parame & "WHERE HIPCUO_NUMOPE ='" & Trim(msk_NumOpe.Text) & "' "
    g_str_Parame = g_str_Parame & "AND HIPCUO_TIPCRO = '" & CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex)) & "' "
    g_str_Parame = g_str_Parame & "AND HIPCUO_NUMCUO = "
    
    'TIPO DE PREPAGO
    If CInt(cmb_TipPre.ItemData(cmb_TipPre.ListIndex)) = 1 Then
        g_str_Parame = g_str_Parame & "     (SELECT (MAX(HIPCUO_NUMCUO)) - (" & r_int_CuoExc & " - " & r_int_NroCuo & " ) "
        g_str_Parame = g_str_Parame & "     FROM CRE_HIPCUO "
        g_str_Parame = g_str_Parame & "     WHERE HIPCUO_NUMOPE ='" & Trim(msk_NumOpe.Text) & "' "
        g_str_Parame = g_str_Parame & "     AND HIPCUO_TIPCRO ='" & CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex)) & "') "
    Else
        If CInt(cmb_TipCro.ItemData(cmb_TipCro.ListIndex)) = 3 Then
            g_str_Parame = g_str_Parame & "     (SELECT HIPMAE_NUMCUO   - (" & r_int_CuoExc & " - " & r_int_NroCuo & " ) "
            g_str_Parame = g_str_Parame & "     FROM CRE_HIPMAE "
            g_str_Parame = g_str_Parame & "     WHERE HIPMAE_NUMOPE ='" & Trim(msk_NumOpe.Text) & "') "
        Else
            g_str_Parame = g_str_Parame & "     (SELECT (HIPMAE_NUMCUO / 6 ) - (" & r_int_CuoExc & " - " & r_int_NroCuo & " ) "
            g_str_Parame = g_str_Parame & "     FROM CRE_HIPMAE "
            g_str_Parame = g_str_Parame & "     WHERE HIPMAE_NUMOPE ='" & Trim(msk_NumOpe.Text) & "') "
        End If
    End If
  
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
        fs_Verificar_CuoPag = False
        Exit Function
    End If
    
    If g_rst_Listas.BOF And g_rst_Listas.EOF Then
        g_rst_Listas.Close
        Set g_rst_Listas = Nothing
        Exit Function
    End If
    
    If g_rst_Listas!HIPCUO_SITUAC = 1 Then
        If MsgBox("Se modificará a partir de la cuota nro " & g_rst_Listas!HIPCUO_NUMCUO & ", pero esta ya ha sido cancelada. ¿Está seguro de seguir con la ejecucción del proceso?", _
                vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
            moddat_g_int_CntErr = 2
        Else
            fs_Verificar_CuoPag = False
        End If
    End If
    
    g_rst_Listas.Close
    Set g_rst_Listas = Nothing
End Function



 
Private Sub msk_NumOpe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call gs_SetFocus(cmd_Buscar)
    End If
End Sub
 

Private Sub txt_NomArc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call gs_SetFocus(cmd_BusArc)
    End If
End Sub

Private Sub txt_NroCuo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call gs_SetFocus(cmd_Import)
    Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
    End If
End Sub

Private Sub cmb_TipCro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call gs_SetFocus(cmb_TipPre)
    End If
End Sub

Private Sub cmb_TipPre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call gs_SetFocus(txt_NomArc)
    End If
End Sub
