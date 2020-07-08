VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Pro_AsgSegInm_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16005
   Icon            =   "OpeTra_frm_347.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   16005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16000
      _Version        =   65536
      _ExtentX        =   28231
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   15915
         _Version        =   65536
         _ExtentX        =   28072
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
            TabIndex        =   2
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
         Begin Threed.SSPanel SSPanel68 
            Height          =   315
            Left            =   690
            TabIndex        =   3
            Top             =   330
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Asignación de Seguro del Inmueble"
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
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_347.frx":000C
            Stretch         =   -1  'True
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7245
         Left            =   30
         TabIndex        =   4
         Top             =   1440
         Width           =   15915
         _Version        =   65536
         _ExtentX        =   28072
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   6855
            Left            =   30
            TabIndex        =   5
            Top             =   330
            Width           =   15795
            _ExtentX        =   27861
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   30
            Cols            =   15
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   3250
            _Version        =   65536
            _ExtentX        =   5733
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
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   4600
            TabIndex        =   7
            Top             =   60
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
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
            Left            =   5910
            TabIndex        =   8
            Top             =   60
            Width           =   3960
            _Version        =   65536
            _ExtentX        =   6985
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   14100
            TabIndex        =   9
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2469
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
               TabIndex        =   16
               Top             =   0
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_Tit_FecGen 
            Height          =   285
            Left            =   9860
            TabIndex        =   17
            Top             =   60
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Generación"
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   11280
            TabIndex        =   18
            Top             =   60
            Width           =   2830
            _Version        =   65536
            _ExtentX        =   4992
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Garantía"
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
            Left            =   3300
            TabIndex        =   10
            Top             =   60
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
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
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   11
         Top             =   750
         Width           =   15915
         _Version        =   65536
         _ExtentX        =   28072
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
         Begin VB.CommandButton cmd_export 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_347.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Exportar datos a excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   660
            Picture         =   "OpeTra_frm_347.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Asignación automática"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Evalua 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_347.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Consulta Operación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   15300
            Picture         =   "OpeTra_frm_347.frx":11F4
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_AsgSegInm_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_dbl_SegInm  As Double
Private Type g_tpo_ActSegInm
   ActSegInm_Col1     As String
   ActSegInm_Col2     As String
   ActSegInm_Col3     As String
End Type

Private Sub cmd_Evalua_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_NomPrd = grd_Listad.Text
   
   grd_Listad.Col = 1
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
   
   grd_Listad.Col = 2
   moddat_g_str_NumOpe = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 2) & Mid(grd_Listad.Text, 8, 5)
   
   grd_Listad.Col = 3
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
   
   grd_Listad.Col = 4
   moddat_g_str_NomCli = grd_Listad.Text
   
   grd_Listad.Col = 5
   moddat_g_str_FecIng = grd_Listad.Text
   
   grd_Listad.Col = 6
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 7
   moddat_g_str_CodSub = grd_Listad.Text
   
   grd_Listad.Col = 9
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgAct = 1
   
   frm_Pro_AsgSegInm_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Inicia
      Call fs_Buscar
      Screen.MousePointer = 0
   Else
      Call fs_Buscar
   End If
End Sub

Private Sub cmd_Export_Click()
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

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_nrofil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_nrofil = 1
   
   With r_obj_Excel.ActiveSheet
      r_int_nrofil = 2
      
      .Cells(r_int_nrofil, 1) = "CREDITOS HIPOTECARIOS - ASIGNACION DE SEGUROS DEL INMUEBLE"
      .Cells(r_int_nrofil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_nrofil, 1).Font.Underline = True
      .Cells(r_int_nrofil, 1).Font.Bold = True
      .Range("A" & r_int_nrofil & ":F" & r_int_nrofil).Merge
      
      r_int_nrofil = 4
      .Cells(r_int_nrofil, 1) = "PRODUCTO":               .Columns("A").ColumnWidth = 38
      .Cells(r_int_nrofil, 2) = "NRO OPERACION":          .Columns("B").ColumnWidth = 18
      .Cells(r_int_nrofil, 3) = "ID CLIENTE":             .Columns("C").ColumnWidth = 18
      .Cells(r_int_nrofil, 4) = "APELLIDOS Y NOMBRES":    .Columns("D").ColumnWidth = 43
      .Cells(r_int_nrofil, 5) = "FECHA GENERACION":       .Columns("E").ColumnWidth = 19
      .Cells(r_int_nrofil, 6) = "TIPO GARANTIA":          .Columns("F").ColumnWidth = 31
      .Cells(r_int_nrofil, 7) = "NOMBRE DEL CONSTRUCTOR": .Columns("G").ColumnWidth = 60
      .Cells(r_int_nrofil, 8) = "NOMBRE DEL PROYECTO":    .Columns("H").ColumnWidth = 49
      .Cells(r_int_nrofil, 9) = "NOMBRE DEL CONSEJERO":   .Columns("I").ColumnWidth = 40
      
      .Columns("A").HorizontalAlignment = xlHAlignLeft
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").HorizontalAlignment = xlHAlignLeft
      
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 9)).Font.Bold = True
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 9)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 9)).HorizontalAlignment = xlHAlignCenter
      
      r_int_nrofil = r_int_nrofil + 1
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_nrofil, 1) = Trim(grd_Listad.TextMatrix(r_int_nroaux, 0))
         .Cells(r_int_nrofil, 2) = Trim(grd_Listad.TextMatrix(r_int_nroaux, 2))
         .Cells(r_int_nrofil, 3) = Trim(grd_Listad.TextMatrix(r_int_nroaux, 3))
         .Cells(r_int_nrofil, 4) = Trim(grd_Listad.TextMatrix(r_int_nroaux, 4))
         .Cells(r_int_nrofil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_nrofil, 6) = Trim(grd_Listad.TextMatrix(r_int_nroaux, 10))
         
         .Cells(r_int_nrofil, 7) = grd_Listad.TextMatrix(r_int_nroaux, 12)
         .Cells(r_int_nrofil, 8) = grd_Listad.TextMatrix(r_int_nroaux, 13)
         .Cells(r_int_nrofil, 9) = grd_Listad.TextMatrix(r_int_nroaux, 14)
         r_int_nrofil = r_int_nrofil + 1
      Next
      
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmd_Proces_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer
Dim r_bol_Estado        As Boolean
Dim r_str_Mensaje       As String
Dim r_int_ConAux        As Integer
Dim r_arr_Matriz()      As g_tpo_ActSegInm

   r_str_Mensaje = ""
   r_bol_Estado = False
   ReDim r_arr_Matriz(0)
   
   'valida selección
   r_int_ConSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 11) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionado Solicitudes para asignación de Seguro del Inmueble.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirma
   If MsgBox("¿Está seguro de Asignar el Seguro del Inmueble a las operaciones seleccionadas?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      
      If (grd_Listad.TextMatrix(r_int_Contad, 11) = "X") Then
         r_dbl_SegInm = 0
         moddat_g_str_NumOpe = Replace(grd_Listad.TextMatrix(r_int_Contad, 2), "-", "")
         ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
         r_arr_Matriz(UBound(r_arr_Matriz)).ActSegInm_Col1 = moddat_g_str_NumOpe
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " SELECT HIPCUO_NUMCUO AS NUMCUO, HIPCUO_FECVCT, HIPCUO_VIVORG "
         g_str_Parame = g_str_Parame & "   FROM CRE_HIPCUO "
         g_str_Parame = g_str_Parame & "  WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
         g_str_Parame = g_str_Parame & "    AND HIPCUO_TIPCRO = 1 "
         g_str_Parame = g_str_Parame & "    AND HIPCUO_FECVCT >= " & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ""
         g_str_Parame = g_str_Parame & "    AND HIPCUO_SITUAC = 2 "
         g_str_Parame = g_str_Parame & "  ORDER BY HIPCUO_NUMCUO "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
              
         If g_rst_Princi.BOF And g_rst_Princi.EOF Then
            g_rst_Princi.Close
            Set g_rst_Princi = Nothing
            Exit Sub
         End If
         
         r_bol_Estado = True
         g_rst_Princi.MoveFirst
         If g_rst_Princi!HIPCUO_VIVORG = 0 Then
            r_dbl_SegInm = moddat_gf_Calcular_SegInm(moddat_g_str_NumOpe)
            r_arr_Matriz(UBound(r_arr_Matriz)).ActSegInm_Col2 = r_dbl_SegInm
            
            Do While Not g_rst_Princi.EOF
               g_str_Parame = " "
               g_str_Parame = g_str_Parame & " UPDATE CRE_HIPCUO SET "
               g_str_Parame = g_str_Parame & " HIPCUO_VIVORG = " & r_dbl_SegInm & ""
               g_str_Parame = g_str_Parame & "  WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "'"
               g_str_Parame = g_str_Parame & "    AND HIPCUO_TIPCRO = 1 "
               g_str_Parame = g_str_Parame & "    AND HIPCUO_NUMCUO = " & CInt(g_rst_Princi!NUMCUO) & " "
               g_str_Parame = g_str_Parame & "    AND HIPCUO_FECVCT >= " & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ""
               g_str_Parame = g_str_Parame & "    AND HIPCUO_SITUAC = 2 "
               
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
                  r_arr_Matriz(UBound(r_arr_Matriz)).ActSegInm_Col3 = "No cargado"
                  GoTo Seguir
               End If
               
               '****REGISTRAR LOG
               If (r_bol_Estado = True) Then
                   r_bol_Estado = False
                   g_str_Parame = ""
                   g_str_Parame = g_str_Parame & "INSERT INTO CRE_SEGINM ("
                   g_str_Parame = g_str_Parame & "SEGINM_NUMOPE, "
                   g_str_Parame = g_str_Parame & "SEGINM_TIPCAR, "
                   g_str_Parame = g_str_Parame & "SEGINM_FECCAR, "
                   g_str_Parame = g_str_Parame & "SEGINM_HORCAR, "
                   g_str_Parame = g_str_Parame & "SEGINM_MTOSEG, "
                   g_str_Parame = g_str_Parame & "SEGINM_CUOCAR, "
                   g_str_Parame = g_str_Parame & "SEGUSUCRE, "
                   g_str_Parame = g_str_Parame & "SEGFECCRE, "
                   g_str_Parame = g_str_Parame & "SEGHORCRE, "
                   g_str_Parame = g_str_Parame & "SEGPLTCRE, "
                   g_str_Parame = g_str_Parame & "SEGTERCRE, "
                   g_str_Parame = g_str_Parame & "SEGSUCCRE) "
                   g_str_Parame = g_str_Parame & "VALUES ( "
                   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
                   g_str_Parame = g_str_Parame & 2 & " , "
                   g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
                   g_str_Parame = g_str_Parame & Format(Time, "HHMMSS") & " , "
                   g_str_Parame = g_str_Parame & r_dbl_SegInm & " , "
                   g_str_Parame = g_str_Parame & CInt(g_rst_Princi!NUMCUO) & " , "
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
                   g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
                   g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & ", "
                   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "' ,"
                   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                                            
                   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                      Exit Sub
                   End If
                   r_arr_Matriz(UBound(r_arr_Matriz)).ActSegInm_Col3 = "Cargado"
                   
               End If

               g_rst_Princi.MoveNext
            Loop
         Else
            g_rst_Princi.Close
            Set g_rst_Princi = Nothing
         End If
      End If
      
Seguir:
   Next r_int_Contad
   
   r_str_Mensaje = ""
   For r_int_ConAux = 1 To UBound(r_arr_Matriz)
      r_str_Mensaje = r_str_Mensaje & IIf(r_int_ConAux = 1, "", Chr(13)) & r_arr_Matriz(r_int_ConAux).ActSegInm_Col1 & " - " & r_arr_Matriz(r_int_ConAux).ActSegInm_Col3 & " - " & r_arr_Matriz(r_int_ConAux).ActSegInm_Col2
   Next r_int_ConAux
   
   MsgBox "Proceso Finalizado." & Chr(13) & Chr(13) & r_str_Mensaje, vbInformation, modgen_g_str_NomPlt

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
   
   Call gs_SetFocus(grd_Listad)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 3250
   grd_Listad.ColWidth(1) = 0
   grd_Listad.ColWidth(2) = 1300
   grd_Listad.ColWidth(3) = 1300
   grd_Listad.ColWidth(4) = 3950 '5135
   grd_Listad.ColWidth(5) = 1420
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 2800
   grd_Listad.ColWidth(11) = 1400       'Seleccionar
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 0
   grd_Listad.ColWidth(14) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignLeftCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_CODPRD, HIPMAE_NUMSOL, HIPMAE_NUMOPE, HIPMAE_TDOCLI, HIPMAE_NDOCLI, "
   g_str_Parame = g_str_Parame & "       HIPMAE_FECACT, HIPMAE_CODPRD, HIPMAE_CODSUB, HIPMAE_MONEDA, PRODUC_DESCRI, "
   g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOM_CLIENTE, "
   g_str_Parame = g_str_Parame & "       P.PARPRD_DESCRI AS MODALIDAD, "
   g_str_Parame = g_str_Parame & "       (SELECT P.PARDES_DESCRI FROM MNT_PARDES P WHERE P.PARDES_CODGRP = 241 AND P.PARDES_CODITE = H.HIPMAE_TIPGAR) TIPOGARANTIA, "
   g_str_Parame = g_str_Parame & "       (SELECT Trim(EJECMC_APEPAT) || ' ' || Trim(EJECMC_APEMAT) || ' ' || Trim(EJECMC_NOMBRE) FROM CRE_EJECMC WHERE EJECMC_CODEJE = HIPMAE_CONHIP) AS NOM_CONSEJERO, "
   g_str_Parame = g_str_Parame & "       (CASE WHEN SOLINM_TABPRY IS NOT NULL THEN "
   g_str_Parame = g_str_Parame & "             CASE WHEN SOLINM_TABPRY = 2 THEN "
   g_str_Parame = g_str_Parame & "                  CASE WHEN SOLINM_PRYCOD IS NOT NULL THEN "
   g_str_Parame = g_str_Parame & "                       CASE WHEN LENGTH (SOLINM_PRYCOD) > 0 THEN "
   g_str_Parame = g_str_Parame & "                            (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = SOLINM_PRYCOD) "
   g_str_Parame = g_str_Parame & "                       Else "
   g_str_Parame = g_str_Parame & "                            CASE WHEN LENGTH (SOLINM_PRYNOM) > 0 THEN TRIM(SOLINM_PRYNOM) END "
   g_str_Parame = g_str_Parame & "                       End "
   g_str_Parame = g_str_Parame & "                  Else "
   g_str_Parame = g_str_Parame & "                       CASE WHEN LENGTH (SOLINM_PRYCOD) > 0 THEN "
   g_str_Parame = g_str_Parame & "                            (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = SOLINM_PRYCOD) "
   g_str_Parame = g_str_Parame & "                       Else "
   g_str_Parame = g_str_Parame & "                            CASE WHEN SOLINM_PRYNOM IS NOT NULL THEN "
   g_str_Parame = g_str_Parame & "                                 Trim (SOLINM_PRYNOM) "
   g_str_Parame = g_str_Parame & "                            Else '' "
   g_str_Parame = g_str_Parame & "                            End "
   g_str_Parame = g_str_Parame & "                       End "
   g_str_Parame = g_str_Parame & "                  End "
   g_str_Parame = g_str_Parame & "             Else "
   g_str_Parame = g_str_Parame & "                 CASE WHEN SOLINM_PRYCOD IS NOT NULL THEN "
   g_str_Parame = g_str_Parame & "                      (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = SOLINM_PRYCOD) "
   g_str_Parame = g_str_Parame & "                 Else "
   g_str_Parame = g_str_Parame & "                      CASE WHEN SOLINM_PRYNOM IS NOT NULL THEN "
   g_str_Parame = g_str_Parame & "                           Trim (SOLINM_PRYNOM) "
   g_str_Parame = g_str_Parame & "                      Else "
   g_str_Parame = g_str_Parame & "                           '' "
   g_str_Parame = g_str_Parame & "                      End "
   g_str_Parame = g_str_Parame & "                 End "
   g_str_Parame = g_str_Parame & "             End "
   g_str_Parame = g_str_Parame & "        Else "
   g_str_Parame = g_str_Parame & "             CASE WHEN SOLINM_PRYCOD IS NOT NULL THEN "
   g_str_Parame = g_str_Parame & "                 (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = SOLINM_PRYCOD) "
   g_str_Parame = g_str_Parame & "             Else "
   g_str_Parame = g_str_Parame & "                 '' "
   g_str_Parame = g_str_Parame & "             End "
   g_str_Parame = g_str_Parame & "        END)  NOM_PROYECTO, "
   g_str_Parame = g_str_Parame & "       NVL(CASE WHEN SOLINM_TIPDOC_CON = 7 THEN TRIM(L.DATGEN_RAZSOC) ELSE TRIM(SOLINM_RAZSOC_CON) END, '') AS NOM_CONSTRUCTOR "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE H "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC ON PRODUC_CODIGO = HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPMAE_TDOCLI AND DATGEN_NUMDOC = HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM ON SOLINM_NUMSOL = H.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN L ON L.DATGEN_EMPTDO = SOLINM_TIPDOC_CON AND L.DATGEN_EMPNDO = SOLINM_NUMDOC_CON "
   g_str_Parame = g_str_Parame & "  LEFT JOIN (SELECT PARPRD_CODPRD, PARPRD_CODSUB, PARPRD_CODITE, PARPRD_DESCRI FROM CRE_PARPRD P "
   g_str_Parame = g_str_Parame & "              WHERE PARPRD_CODGRP = '003') P ON "
   g_str_Parame = g_str_Parame & "                    P.PARPRD_CODPRD = H.HIPMAE_CODPRD AND "
   g_str_Parame = g_str_Parame & "                    P.PARPRD_CODITE = TRIM(to_char(H.HIPMAE_CODMOD, '000')) AND "
   g_str_Parame = g_str_Parame & "                    P.PARPRD_CODSUB = H.HIPMAE_CODSUB "
   g_str_Parame = g_str_Parame & "  LEFT JOIN (SELECT HIPCUO_NUMOPE, SUM(HIPCUO_VIVORG) SEGINM FROM CRE_HIPCUO HI "
   g_str_Parame = g_str_Parame & "              WHERE HIPCUO_TIPCRO = 1 GROUP BY HIPCUO_NUMOPE) HI ON H.HIPMAE_NUMOPE = HI.HIPCUO_NUMOPE "
   g_str_Parame = g_str_Parame & " WHERE (HIPMAE_SITUAC = 1 OR HIPMAE_SITUAC = 2) "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_FECDES >= 20150801 "
   g_str_Parame = g_str_Parame & "   AND HI.SEGINM = 0 "
   g_str_Parame = g_str_Parame & "   AND P.PARPRD_DESCRI LIKE 'BIEN FUTURO%' "
   g_str_Parame = g_str_Parame & " ORDER BY HIPMAE_FECACT DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = Trim(g_rst_Princi!PRODUC_DESCRI)
         
         grd_Listad.Col = 1
         grd_Listad.Text = Left(g_rst_Princi!HIPMAE_NUMSOL, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMSOL, 4, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMSOL, 7, 2) & "-" & Right(g_rst_Princi!HIPMAE_NUMSOL, 4)
         
         grd_Listad.Col = 2
         grd_Listad.Text = Left(g_rst_Princi!hipmae_numope, 3) & "-" & Mid(g_rst_Princi!hipmae_numope, 4, 2) & "-" & Right(g_rst_Princi!hipmae_numope, 5)
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
         
         grd_Listad.Col = 4
         grd_Listad.Text = g_rst_Princi!NOM_CLIENTE
         
         grd_Listad.Col = 5
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT))
         
         grd_Listad.Col = 6
         grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_CODPRD & "")
         
         grd_Listad.Col = 7
         grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_CODSUB & "")
         
         grd_Listad.Col = 8
         grd_Listad.Text = g_rst_Princi!HIPMAE_FECACT
         
         grd_Listad.Col = 9
         grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_MONEDA)
         
         grd_Listad.Col = 10
         grd_Listad.Text = CStr(g_rst_Princi!TIPOGARANTIA)
         '-----
         grd_Listad.Col = 12
         grd_Listad.Text = CStr(g_rst_Princi!NOM_CONSTRUCTOR)
         
         grd_Listad.Col = 13
         grd_Listad.Text = CStr(g_rst_Princi!NOM_PROYECTO)
         
         grd_Listad.Col = 14
         grd_Listad.Text = CStr(g_rst_Princi!NOM_CONSEJERO)
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   If grd_Listad.Rows = 0 Then
      cmd_Evalua.Enabled = False
      cmd_Proces.Enabled = False
      MsgBox "No se encontraron Operaciones para Asignar Seguro del Inmueble.", vbInformation, modgen_g_str_NomPlt
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Function moddat_gf_Calcular_SegInm(ByVal p_NumOpe As String) As String
   moddat_gf_Calcular_SegInm = ""
   
   'Cálculo del Seguro del Inmueble
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT ROUND((EVATAS_SUMASE_INM+EVATAS_SUMASE_ES1+EVATAS_SUMASE_ES2+EVATAS_SUMASE_DEP)*((SELECT HIPMAE_FOIVIV FROM CRE_HIPMAE WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "')/100), 2) SEGINM "
   g_str_Parame = g_str_Parame & "   FROM TRA_EVATAS "
   g_str_Parame = g_str_Parame & "  WHERE EVATAS_NUMSOL = (SELECT HIPMAE_NUMSOL FROM CRE_HIPMAE WHERE HIPMAE_NUMOPE = '" & p_NumOpe & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   g_rst_Genera.MoveFirst
   moddat_gf_Calcular_SegInm = g_rst_Genera!SEGINM
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 11) = ""
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 11) = "X"
         Next r_Fila
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_Click()
  If grd_Listad.Rows > 0 Then
      If grd_Listad.TextMatrix(grd_Listad.Row, 11) = "X" Then
         grd_Listad.TextMatrix(grd_Listad.Row, 11) = ""
      Else
         grd_Listad.TextMatrix(grd_Listad.Row, 11) = "X"
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Evalua_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecGen_Click()
   If Len(Trim(pnl_Tit_FecGen.Tag)) = 0 Or pnl_Tit_FecGen.Tag = "D" Then
      pnl_Tit_FecGen.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "N")
   Else
      pnl_Tit_FecGen.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Len(Trim(pnl_Tit_NumSol.Tag)) = 0 Or pnl_Tit_NumSol.Tag = "D" Then
      pnl_Tit_NumSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 10, "C")
   Else
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 10, "C-")
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

