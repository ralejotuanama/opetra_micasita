VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_LevCon_08 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10245
   ClientLeft      =   2085
   ClientTop       =   765
   ClientWidth     =   11640
   Icon            =   "OpeTra_frm_112.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   18071
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
         Height          =   3465
         Left            =   30
         TabIndex        =   1
         Top             =   4800
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6112
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
         Begin TabDlg.SSTab tab_Seguim 
            Height          =   3345
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   5900
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento de Instancias"
            TabPicture(0)   =   "OpeTra_frm_112.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label7"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label11"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "pnl_DesOcu"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel5"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel14"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel13"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "grd_LisOcu"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel10"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txt_Observ"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txt_Descar"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "OpeTra_frm_112.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label4"
            Tab(1).Control(1)=   "Label3"
            Tab(1).Control(2)=   "Label6"
            Tab(1).Control(3)=   "pnl_TipAut"
            Tab(1).Control(4)=   "pnl_DesExc"
            Tab(1).Control(5)=   "SSPanel16"
            Tab(1).Control(6)=   "SSPanel15"
            Tab(1).Control(7)=   "SSPanel12"
            Tab(1).Control(8)=   "SSPanel11"
            Tab(1).Control(9)=   "SSPanel9"
            Tab(1).Control(10)=   "grd_LisExc"
            Tab(1).Control(11)=   "txt_ObsExc"
            Tab(1).ControlCount=   12
            TabCaption(2)   =   "Aprobaci�n Condicionada"
            TabPicture(2)   =   "OpeTra_frm_112.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label12"
            Tab(2).Control(1)=   "Label14"
            Tab(2).Control(2)=   "Label15"
            Tab(2).Control(3)=   "pnl_InsCon"
            Tab(2).Control(4)=   "SSPanel20"
            Tab(2).Control(5)=   "SSPanel19"
            Tab(2).Control(6)=   "SSPanel18"
            Tab(2).Control(7)=   "grd_LisCon"
            Tab(2).Control(8)=   "SSPanel17"
            Tab(2).Control(9)=   "txt_LevCon"
            Tab(2).Control(10)=   "txt_ObsCon"
            Tab(2).ControlCount=   11
            Begin VB.TextBox txt_Descar 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Text            =   "OpeTra_frm_112.frx":0060
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_Observ 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Text            =   "OpeTra_frm_112.frx":0064
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   975
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Text            =   "OpeTra_frm_112.frx":0068
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Text            =   "OpeTra_frm_112.frx":006C
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "OpeTra_frm_112.frx":0070
               Top             =   1980
               Width           =   10005
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   45
               Left            =   30
               TabIndex        =   8
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisOcu 
               Height          =   855
               Left            =   30
               TabIndex        =   9
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   285
               Left            =   60
               TabIndex        =   10
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Ocurrencia"
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   2400
               TabIndex        =   11
               Top             =   360
               Width           =   8595
               _Version        =   65536
               _ExtentX        =   15161
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripci�n Ocurrencia"
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
               Left            =   1230
               TabIndex        =   12
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Ocurrencia"
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
            Begin Threed.SSPanel pnl_DesOcu 
               Height          =   315
               Left            =   1320
               TabIndex        =   13
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "D�a: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisExc 
               Height          =   855
               Left            =   -74970
               TabIndex        =   14
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -74940
               TabIndex        =   15
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Excepci�n"
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -69330
               TabIndex        =   16
               Top             =   360
               Width           =   5325
               _Version        =   65536
               _ExtentX        =   9393
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripci�n Excepci�n"
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   285
               Left            =   -73770
               TabIndex        =   17
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Excepci�n"
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -72600
               TabIndex        =   18
               Top             =   360
               Width           =   3285
               _Version        =   65536
               _ExtentX        =   5794
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
            Begin Threed.SSPanel SSPanel16 
               Height          =   45
               Left            =   -74970
               TabIndex        =   19
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            End
            Begin Threed.SSPanel pnl_DesExc 
               Height          =   315
               Left            =   -73680
               TabIndex        =   20
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "D�a: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_TipAut 
               Height          =   315
               Left            =   -73650
               TabIndex        =   21
               Top             =   2970
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "D�a: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel SSPanel17 
               Height          =   45
               Left            =   -74970
               TabIndex        =   22
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisCon 
               Height          =   855
               Left            =   -74970
               TabIndex        =   23
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   4
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -74940
               TabIndex        =   24
               Top             =   360
               Width           =   2745
               _Version        =   65536
               _ExtentX        =   4842
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
            Begin Threed.SSPanel SSPanel19 
               Height          =   285
               Left            =   -65610
               TabIndex        =   25
               Top             =   360
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Situaci�n"
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   285
               Left            =   -72210
               TabIndex        =   26
               Top             =   360
               Width           =   6615
               _Version        =   65536
               _ExtentX        =   11668
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Condiciones de Aprobaci�n"
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
            Begin Threed.SSPanel pnl_InsCon 
               Height          =   315
               Left            =   -73680
               TabIndex        =   27
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "D�a: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   60
               TabIndex        =   36
               Top             =   2640
               Width           =   1035
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   60
               TabIndex        =   35
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observaci�n:"
               Height          =   495
               Left            =   60
               TabIndex        =   34
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripci�n:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   33
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label3 
               Caption         =   "Excepci�n:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   32
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label6 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   31
               Top             =   2970
               Width           =   1095
            End
            Begin VB.Label Label12 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   615
               Left            =   -74940
               TabIndex        =   30
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   29
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobaci�n:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   28
               Top             =   1980
               Width           =   1155
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   37
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   7830
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
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   7260
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin Threed.SSPanel pnl_AprCon 
            Height          =   555
            Left            =   8460
            TabIndex        =   38
            Top             =   60
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "CLIENTE CON APROBACION CONDICIONADA PENDIENTE"
            ForeColor       =   16777215
            BackColor       =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   600
            TabIndex        =   64
            Top             =   60
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Levantamiento de Aprobaci�n Condicionada"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   600
            TabIndex        =   65
            Top             =   360
            Width           =   5055
            _Version        =   65536
            _ExtentX        =   8916
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tr�mites COFIDE"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            Picture         =   "OpeTra_frm_112.frx":0074
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel36 
         Height          =   645
         Left            =   30
         TabIndex        =   39
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_LevCon 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_112.frx":037E
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Levantar Condici�n de Aprobaci�n"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_112.frx":07C0
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   41
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   42
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   43
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
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
         End
         Begin Threed.SSPanel pnl_FecSol 
            Height          =   315
            Left            =   9450
            TabIndex        =   44
            Top             =   60
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "01/01/9999"
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
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   47
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   46
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   8070
            TabIndex        =   45
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1875
         Left            =   30
         TabIndex        =   48
         Top             =   8310
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3307
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
         Begin VB.TextBox txt_ObsEva 
            Height          =   705
            Left            =   30
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Text            =   "OpeTra_frm_112.frx":0C02
            Top             =   1140
            Width           =   11445
         End
         Begin MSFlexGridLib.MSFlexGrid grd_LisEva 
            Height          =   1065
            Left            =   30
            TabIndex        =   50
            Top             =   60
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   1879
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2505
         Left            =   30
         TabIndex        =   51
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   4419
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   2385
            Left            =   60
            TabIndex        =   52
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   4207
            _Version        =   393216
            Style           =   1
            Tabs            =   10
            TabsPerRow      =   10
            TabHeight       =   520
            TabCaption(0)   =   "Datos Cliente"
            TabPicture(0)   =   "OpeTra_frm_112.frx":0C06
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos C�nyuge"
            TabPicture(1)   =   "OpeTra_frm_112.frx":0C22
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Inmueble"
            TabPicture(2)   =   "OpeTra_frm_112.frx":0C3E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Datos Cr�dito"
            TabPicture(3)   =   "OpeTra_frm_112.frx":0C5A
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Ev. Crediticia"
            TabPicture(4)   =   "OpeTra_frm_112.frx":0C76
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(4)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Tasaci�n"
            TabPicture(5)   =   "OpeTra_frm_112.frx":0C92
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(5)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Ev. Seguros"
            TabPicture(6)   =   "OpeTra_frm_112.frx":0CAE
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "grd_Listad(6)"
            Tab(6).ControlCount=   1
            TabCaption(7)   =   "Informe Legal"
            TabPicture(7)   =   "OpeTra_frm_112.frx":0CCA
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "txt_InfLeg"
            Tab(7).ControlCount=   1
            TabCaption(8)   =   "Contratos"
            TabPicture(8)   =   "OpeTra_frm_112.frx":0CE6
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "grd_Listad(7)"
            Tab(8).ControlCount=   1
            TabCaption(9)   =   "Bloqueo Registral"
            TabPicture(9)   =   "OpeTra_frm_112.frx":0D02
            Tab(9).ControlEnabled=   0   'False
            Tab(9).Control(0)=   "grd_Listad(8)"
            Tab(9).ControlCount=   1
            Begin VB.TextBox txt_InfLeg 
               Height          =   1965
               Left            =   -74940
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   53
               Text            =   "OpeTra_frm_112.frx":0D1E
               Top             =   360
               Width           =   11235
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1965
               Index           =   0
               Left            =   60
               TabIndex        =   54
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3466
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1965
               Index           =   1
               Left            =   -74940
               TabIndex        =   55
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3466
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1965
               Index           =   2
               Left            =   -74940
               TabIndex        =   56
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3466
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1965
               Index           =   3
               Left            =   -74940
               TabIndex        =   57
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3466
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1965
               Index           =   4
               Left            =   -74940
               TabIndex        =   58
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3466
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1965
               Index           =   5
               Left            =   -74940
               TabIndex        =   59
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3466
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1965
               Index           =   6
               Left            =   -74940
               TabIndex        =   60
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3466
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1965
               Index           =   7
               Left            =   -74940
               TabIndex        =   61
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3466
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1965
               Index           =   8
               Left            =   -74940
               TabIndex        =   62
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3466
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
         End
      End
   End
End
Attribute VB_Name = "frm_LevCon_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_AprCon     As Integer

Private Sub cmd_LevCon_Click()
   moddat_g_int_CodIns = 62
   moddat_g_str_Observ = ""
   
   moddat_g_int_FlgAct_1 = 1
   
   frm_RecSol_17.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      
      If Not moddat_gf_Inserta_LevCon(moddat_g_str_NumSol, 61, moddat_g_str_Observ) Then
         Exit Sub
      End If
      
      Screen.MousePointer = 0
      
      'Enviando Correo Electr�nico
      modgen_g_str_Mail_Asunto = "TRAMITES COFIDE - LEVANTAMIENTO DE CONDICIONES DE APROBACION (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & " <br>"
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " <br>"
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & " <br>"
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & " <br>"
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ & " <br>"
      
      Call fs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
      
      MsgBox "Se levanto las condiciones de la Aprobaci�n.", vbInformation, modgen_g_str_NomPlt
      
      moddat_g_int_FlgAct = 2
      
      Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   moddat_g_int_CodIns = 62
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   
   'Buscar Informaci�n de Solicitud de Cr�dito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
 
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)       'Buscar Informaci�n del Cliente
   Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)       'Buscar Informaci�n del C�nyuge
   
   Call fs_DatInm          'Datos del Inmueble
   Call fs_DatCre          'Datos del Cr�dito
   
   Call fs_EvaCre          'Datos de Evaluaci�n Crediticia
   Call fs_DatTas          'Datos de Tasaci�n
   Call fs_DatSeg          'Datos de Seguros
   Call fs_DatLeg          'Datos de Legal
   
   Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
   Call fs_Buscar_LisExc      'Buscando Excepciones
   Call fs_Buscar_LisCon      'Buscando Aprobaciones Condicionadas
   
   'Si no hay Excepciones aplicadas
   If grd_LisExc.Rows = 0 Then
      tab_Seguim.TabVisible(1) = False
   End If

   'Si no hay Aprobaciones Condicionadas
   If grd_LisCon.Rows = 0 Then
      tab_Seguim.TabVisible(2) = False
   End If
   
   'Si no hay Aprobaciones Condicionadas Pendiente
   If l_int_AprCon = 0 Then
      pnl_AprCon.Visible = False
   End If
   
   'Cargando Datos de Evaluaci�n Crediticia
   Call fs_Buscar_DatEva
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer

   'Inicializando Grid de Cliente y de C�nyuge
   For r_int_Contad = 0 To 8
      grd_Listad(r_int_Contad).ColWidth(0) = 2900
      grd_Listad(r_int_Contad).ColWidth(1) = 7950
   
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad

   'Lista de Ocurrencias
   grd_LisOcu.ColWidth(0) = 1155
   grd_LisOcu.ColWidth(1) = 1185
   grd_LisOcu.ColWidth(2) = 8595
   grd_LisOcu.ColWidth(3) = 0
   grd_LisOcu.ColWidth(4) = 0
   
   grd_LisOcu.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOcu.ColAlignment(1) = flexAlignCenterCenter
   grd_LisOcu.ColAlignment(2) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_LisOcu)

   pnl_DesOcu.Caption = ""
   txt_Observ.Text = ""
   txt_Descar.Text = ""

   'Lista de Excepciones
   grd_LisExc.ColWidth(0) = 1175
   grd_LisExc.ColWidth(1) = 1175
   grd_LisExc.ColWidth(2) = 3275
   grd_LisExc.ColWidth(3) = 5325
   grd_LisExc.ColWidth(4) = 0
   
   grd_LisExc.ColAlignment(0) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(1) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(2) = flexAlignLeftCenter
   grd_LisExc.ColAlignment(3) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_LisExc)

   pnl_DesExc.Caption = ""
   txt_ObsExc.Text = ""
   pnl_TipAut.Caption = ""

   'Lista de Aprobaciones Condicionadas
   grd_LisCon.ColWidth(0) = 2735
   grd_LisCon.ColWidth(1) = 6605
   grd_LisCon.ColWidth(2) = 1625
   grd_LisCon.ColWidth(3) = 0
   
   grd_LisCon.ColAlignment(0) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(1) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(2) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_LisCon)

   pnl_InsCon.Caption = ""
   txt_ObsCon.Text = ""
   txt_LevCon.Text = ""

   'Lista de Datos de Evaluaci�n
   grd_LisEva.ColWidth(0) = 3300
   grd_LisEva.ColWidth(1) = 7940

   grd_LisEva.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEva.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Documento de Identidad"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Apellidos y Nombres"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
      
      If g_rst_Princi!DatGen_FLGDOA = 1 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Documento Adicional de Identidad"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DatGen_FLGDOA)) & IIf(g_rst_Princi!DatGen_FLGDOA = 1, " ( " & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TIPDOA)) & " - " & Trim(g_rst_Princi!DatGen_NUMDOA) & ")", "")
      End If
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Sexo"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("207", CStr(g_rst_Princi!DatGen_CodSex))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Fecha de Nacimiento"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Nacionalidad"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Lugar de Nacimiento"
   
      If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      Else
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = "<< NO REGISTRADO >>"
      End If
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Estado Civil"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DATGEN_REGCYG), "")
         
         If g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
            moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
         End If
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Nivel de Estudios"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Profesi�n"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DatGen_Profes))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Celular"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Nro. Dependientes Econ�micos"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = CStr(g_rst_Princi!DatGen_DepEco)
      
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Edades"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = IIf(g_rst_Princi!DatGen_EDAD01 > 0, CStr(g_rst_Princi!DatGen_EDAD01), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD02 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD02), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD03 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD03), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD04 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD04), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD05 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD05), "")
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "E-mail"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_DirEle & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Autorizaci�n Env�o"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_AUTENV))
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Domicilio"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Tel�fono Domicilio"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      End If
      
      If p_Indice = 0 Then
         moddat_g_str_FecNac_Tit = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      Else
         moddat_g_str_FecNac_Cyg = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      End If
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Actividad Econ�mica Principal
   Call fs_ActEco(p_TipDoc, p_NumDoc, 1, p_Indice)
   Call fs_ActEco(p_TipDoc, p_NumDoc, 2, p_Indice)
End Sub

Private Sub fs_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer, ByVal p_Indice As Integer)
   Dim r_var_ColTxt
   
   If p_OrdAct = 1 Then
      r_var_ColTxt = modgen_g_con_ColAzu
   Else
      r_var_ColTxt = modgen_g_con_ColRoj
   End If

   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 2
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = r_var_ColTxt
      grd_Listad(p_Indice).Text = "Ocupaci�n " & Left(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct)), 1) & Mid(LCase(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct))), 2)
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = r_var_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("008", g_rst_Princi!ACTECO_CODACT)
      
      Select Case g_rst_Princi!ACTECO_CODACT
         Case 11: Call fs_ActEco_Dep(p_Indice, r_var_ColTxt)
         Case 21: Call fs_ActEco_Ind(p_Indice, r_var_ColTxt)
         Case 31: Call fs_ActEco_Com(p_Indice, r_var_ColTxt)
         Case 41: Call fs_ActEco_Acc(p_Indice, r_var_ColTxt)
         Case 51: Call fs_ActEco_Ren(p_Indice, r_var_ColTxt)
         Case 61: Call fs_ActEco_Otr(p_Indice, r_var_ColTxt)
      End Select
      
      If p_Indice = 0 And p_OrdAct = 1 Then
         moddat_g_int_FlgActPri_Cli = 1
      End If
      
      If p_Indice = 0 And p_OrdAct = 2 Then
         moddat_g_int_FlgActSec_Cli = 1
      End If
      
      If p_Indice = 1 And p_OrdAct = 1 Then
         moddat_g_int_FlgActPri_Cyg = 1
      End If
      
      If p_Indice = 1 And p_OrdAct = 2 Then
         moddat_g_int_FlgActSec_Cyg = 1
      End If
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_ActEco_Dep(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad Empleador"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Situaci�n como Trabajador"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("235", g_rst_Princi!ActEco_Dep_SitTra)

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_RazSoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomCom & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Dep_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Dep_CodCiu))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TeleRH & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_AnexRH & "")) > 0, " ANEXO: " & Trim(g_rst_Princi!ActEco_Dep_AnexRH & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Direcci�n"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
   Else
      g_rst_Genera.MoveFirst

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELERH & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELERH & "")) > 0, " ANEXO: " & Trim(g_rst_Genera!DATGEN_ANEXRH & ""), "")

      If g_rst_Princi!ActEco_Dep_TipOfi = 1 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Direcci�n"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                     IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_REFERE & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Fax"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
      Else
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Direcci�n"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                     " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                     IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Fax"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
      End If
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Dep_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Frecuencia de Haberes"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("210", CStr(g_rst_Princi!ActEco_Dep_FreHab))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Ingreso"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Cargo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Dep_CodCar = "999999", Trim(g_rst_Princi!ActEco_Dep_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Dep_CodCar))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Area"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomAre & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Anexo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tel�fono Directo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Celular"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Celula & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "E-mail"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_DirEle & "")

   If g_rst_Princi!ActEco_Dep_TraAnt = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Documento Identidad Empleador Anterior"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc_Ant) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc_Ant & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc_Ant) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc_Ant & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social (Empleador Anterior)"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_RazSoc_Ant & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Nombre Comercial (Empleador Anterior)"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomCom_Ant & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s) (Empleador Anterior)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1_Ant & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Ingreso (Empleador Anterior)"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng_Ant))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Cese (Empleador Anterior)"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecCes_Ant))
   End If
End Sub

Private Sub fs_ActEco_Ind(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Direcci�n"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Ind_TipVia)) & _
                               " " & Trim(g_rst_Princi!ActEco_Ind_NomVia) & " " & Trim(g_rst_Princi!ActEco_Ind_NumVia) & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Ind_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Ind_IntDpt) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Ind_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Ind_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Ind_NomZon), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Referencia"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Refere & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Indartamento / Provincia / Distrito"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ind_UbiGeo))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tel�fono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2 & ""), "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fax"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Ind_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Ind_CodCiu))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ind_IngNet, 15, 2)
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Actividades"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Contrato de Locaci�n de Servicios"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
   
   If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Documento Identidad Empleador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Ind_NumDoc_Emp & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_RazSoc_Emp & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Nombre Comercial"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_NomCom_Emp & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1_Emp & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Raz�n Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Tel�fono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Ingreso"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Cargo"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Ind_CodCar = "999999", Trim(g_rst_Princi!ActEco_Ind_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Ind_CodCar))
   End If
End Sub

Private Sub fs_ActEco_Com(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Com_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Com_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Raz�n Social"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Nombre Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Direcci�n"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Com_TipVia)) & _
                               " " & Trim(g_rst_Princi!ActEco_Com_NomVia) & " " & Trim(g_rst_Princi!ActEco_Com_NumVia) & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Com_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Com_IntDpt) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Com_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Com_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Com_NomZon), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Referencia"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_Refere & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Com_UbiGeo))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tel�fono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Com_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Com_Telef2 & ""), "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fax"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NumFax & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Com_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Com_CodCiu))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Giro Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_GirCom(g_rst_Princi!ActEco_Com_GirCom)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ventas Mensuales"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_VtaMen, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Operaciones"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Cargo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Com_CodCar = "999999", Trim(g_rst_Princi!ActEco_Com_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Com_CodCar))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "R�gimen Tributario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("215", CStr(g_rst_Princi!ActEco_Com_RegTri))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participaci�n"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_PorPar, 7, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tipo de Local Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("208", CStr(g_rst_Princi!ActEco_Com_TipLoc))
   
   If g_rst_Princi!ActEco_Com_TipLoc = 2 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Alquiler Mensual"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_AlqMen, 15, 2)
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Arrendador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NomArr & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono Arrendador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_TelArr & "")
   End If
End Sub

Private Sub fs_ActEco_Acc(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Acc_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Acc_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Acc_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_RazSoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_NomCom & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Acc_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Acc_CodCiu))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Direcci�n"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Acc_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Acc_NomVia) & " " & Trim(g_rst_Princi!ActEco_Acc_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Acc_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Acc_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Acc_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Acc_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Acc_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Acc_UbiGeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Acc_UbiGeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Acc_UbiGeo))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Acc_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_NumFax & "")
   Else
      g_rst_Genera.MoveFirst

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Raz�n Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Direcci�n"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                  " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                  IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_REFERE & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Antig�edad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participaci�n"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_PorPar, 7, 2)
End Sub

Private Sub fs_ActEco_Ren(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Direcci�n de Propiedad 01"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Nombre de Arrendatario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Alquiler"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tel�fono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele21 & ""), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Alquiler Mensual"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe1, 15, 2)
   
   If g_rst_Princi!ActEco_Ren_SegPro = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Direcci�n de Propiedad 02"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre de Arrendatario"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Inicio de Alquiler"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Tel�fono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele22 & ""), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Alquiler Mensual"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe2, 15, 2)
   End If
End Sub

Private Sub fs_ActEco_Otr(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Otr_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Actividad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Otr_Activi & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Otr_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Otr_CodCiu))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Observaciones"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Otr_Observ & "")
End Sub

Private Sub fs_DatInm()
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(2).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Modalidad"
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLINM_CODMOD)), "000")) Then
         grd_Listad(2).Col = 1
         grd_Listad(2).Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Tipo de Inmueble"
         
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Direcci�n"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Referencia"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Estacionamiento"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Departamento / Provincia / Distrito"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 2
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Proyecto miCasita"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         If Not IsNull(g_rst_Princi!SOLINM_PRYBCO) Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = "Proyecto anclado en Otra IFI"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
         End If
         
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = "Nombre Proyecto"
   
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Nombre Proyecto"
   
               grd_Listad(2).Col = 1
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            End If
         End If
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 2
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Propietario / Promotor"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Docum. Identidad Propietario/Promotor"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Nombre o Raz�n Social"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Direcci�n"
         
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Referencia"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Tel�fono"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         
         If g_rst_Princi!SOLINM_FLGCON = 1 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Docum. Identidad Constructor"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Nombre o Raz�n Social"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Direcci�n"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Referencia"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Tel�fono"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
         End If
      Else
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
            If g_rst_Princi!SOLINM_PRYMCS = 1 Then
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Proyecto Vinculado"
            Else
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Entidad Financiera"
         
               grd_Listad(2).Col = 1
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
               
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Proyecto No Vinculado"
            End If
         
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Or CInt(g_rst_Princi!SOLINM_CODMOD) = 4 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Docum. Identidad Propietario"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Nombre o Raz�n Social"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Direcci�n"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Referencia"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Tel�fono"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         Else
            'Promotor
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Doc. Ident. Promotor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Raz�n Social Promotor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            'Constructor
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Doc. Ident. Constructor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_CON) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_CON)
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Raz�n Social Constructor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_CON, g_rst_Princi!SOLINM_NUMDOC_CON)
         End If
      End If
      
      grd_Listad(2).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(2))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   Call gs_LimpiaGrid(grd_Listad(3))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   grd_Listad(3).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(3).Rows = grd_Listad(3).Rows + 1
   grd_Listad(3).Row = grd_Listad(3).Rows - 1
   grd_Listad(3).Col = 0
   grd_Listad(3).Text = "Producto"

   grd_Listad(3).Col = 1
   grd_Listad(3).Text = moddat_g_str_NomPrd
   
   grd_Listad(3).Rows = grd_Listad(3).Rows + 1
   grd_Listad(3).Row = grd_Listad(3).Rows - 1
   grd_Listad(3).Col = 0
   grd_Listad(3).Text = "Sub-Producto"

   grd_Listad(3).Col = 1
   grd_Listad(3).Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   grd_Listad(3).Rows = grd_Listad(3).Rows + 1
   grd_Listad(3).Row = grd_Listad(3).Rows - 1
   grd_Listad(3).Col = 0
   grd_Listad(3).Text = "Tipo de Evaluaci�n"

   grd_Listad(3).Col = 1
   grd_Listad(3).Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
   
   grd_Listad(3).Rows = grd_Listad(3).Rows + 1
   grd_Listad(3).Row = grd_Listad(3).Rows - 1
   grd_Listad(3).Col = 0
   grd_Listad(3).Text = "Moneda del Pr�stamo"

   grd_Listad(3).Col = 1
   grd_Listad(3).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   grd_Listad(3).Rows = grd_Listad(3).Rows + 1
   grd_Listad(3).Row = grd_Listad(3).Rows - 1
   grd_Listad(3).Col = 0
   grd_Listad(3).Text = "Fecha de Solicitud"

   grd_Listad(3).Col = 1
   grd_Listad(3).Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))

   If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         grd_Listad(3).Rows = grd_Listad(3).Rows + 2
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Valor de Compra Venta"
      
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Aporte Propio"
      
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Monto Pr�stamo"
      
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
      Else
         grd_Listad(3).Rows = grd_Listad(3).Rows + 2
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Valor de Compra Venta"
      
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Aporte Propio"
      
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Monto Pr�stamo"
      
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
      End If
   
      grd_Listad(3).Rows = grd_Listad(3).Rows + 2
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Tasa de Inter�s"
   
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
   
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Plazo"
   
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = CStr(g_rst_Princi!SOLMAE_PLAANO) & " A�os"
   
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Per�odo de Gracia (Meses)"
   
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = CStr(g_rst_Princi!SOLMAE_PERGRA) & " Meses"
   
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Cuotas Extraordinarias"
   
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_CUOEXT))
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Compa��a de Seguros"
   
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Tipo de Seguro Desgravamen"
   
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "D�a de Pago"
   
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   End If
   
   If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
      grd_Listad(3).Rows = grd_Listad(3).Rows + 2
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Instituci�n Financiera de Ahorro"
   
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Monto M�nimo de Ahorro Mensual"
   
      grd_Listad(3).Col = 1
      grd_Listad(3).CellFontName = "Lucida Console"
      grd_Listad(3).CellFontSize = 8
      grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_MONAHO)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
   
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Meses Ahorrados"
   
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = CStr(g_rst_Princi!SOLMAE_MESAHO)
   End If
   
   grd_Listad(3).Rows = grd_Listad(3).Rows + 2
   grd_Listad(3).Row = grd_Listad(3).Rows - 1
   grd_Listad(3).Col = 0
   grd_Listad(3).Text = "Consejero Hipotecario"

   grd_Listad(3).Col = 1
   grd_Listad(3).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
   
   grd_Listad(3).Rows = grd_Listad(3).Rows + 1
   grd_Listad(3).Row = grd_Listad(3).Rows - 1
   grd_Listad(3).Col = 0
   grd_Listad(3).Text = "Ejecutivo de Seguimiento"

   grd_Listad(3).Col = 1
   grd_Listad(3).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)
   
   grd_Listad(3).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(3))
   
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
   moddat_g_str_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_LisOcu()
   Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisOcu)
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 62 "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisOcu.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisOcu.Rows = grd_LisOcu.Rows + 1
      grd_LisOcu.Row = grd_LisOcu.Rows - 1
      
      'Fecha de Ocurrencia
      grd_LisOcu.Col = 0
      grd_LisOcu.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Ocurrencia
      grd_LisOcu.Col = 1
      grd_LisOcu.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Descripci�n Ocurrencia
      grd_LisOcu.Col = 2
      grd_LisOcu.Text = moddat_gf_Consulta_ParDes("004", Format(g_rst_Princi!SEGDET_CODOCU, "000000"))
      
      If g_rst_Princi!SEGFECACT > 0 Then
         r_str_FecOcu = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
         
         grd_LisOcu.Text = grd_LisOcu.Text & " (DESCARGO EFECTUADO - " & r_str_FecOcu
         grd_LisOcu.Text = grd_LisOcu.Text & " / " & gf_FormatoHora(Format(g_rst_Princi!SEGHORACT, "000000")) & ")"
      End If
      
      grd_LisOcu.Col = 3
      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
      
      grd_LisOcu.Col = 4
      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSDES & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisOcu.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisOcu)
   Call grd_LisOcu_Click
End Sub

Private Sub fs_Buscar_LisExc()
   Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisExc)
   
   g_str_Parame = "SELECT * FROM TRA_SEGEXC WHERE "
   g_str_Parame = g_str_Parame & "SEGEXC_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisExc.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisExc.Rows = grd_LisExc.Rows + 1
      grd_LisExc.Row = grd_LisExc.Rows - 1
      
      'Fecha de Excepci�n
      grd_LisExc.Col = 0
      grd_LisExc.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Excepci�n
      grd_LisExc.Col = 1
      grd_LisExc.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Instancia
      grd_LisExc.Col = 2
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGEXC_CODINS))
      
      'Descripci�n Excepci�n
      grd_LisExc.Col = 3
      grd_LisExc.Text = Trim(g_rst_Princi!SEGEXC_DESCRI & "")
      
      'Tipo Autorizaci�n
      grd_LisExc.Col = 4
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("243", CStr(g_rst_Princi!SEGEXC_TIPAUT))
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisExc.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisExc)
   Call grd_LisExc_Click
End Sub

Private Sub fs_Buscar_LisCon()
   l_int_AprCon = 0
   
   Call gs_LimpiaGrid(grd_LisCon)
   
   g_str_Parame = "SELECT * FROM TRA_SEGCON WHERE "
   g_str_Parame = g_str_Parame & "SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGCON_SITUAC ASC, SEGCON_CODINS DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisCon.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisCon.Rows = grd_LisCon.Rows + 1
      grd_LisCon.Row = grd_LisCon.Rows - 1
      
      'Instancia
      grd_LisCon.Col = 0
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGCON_CODINS))
      
      'Descripci�n Condiciones
      grd_LisCon.Col = 1
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSCON & "")
      
      'Situaci�n
      grd_LisCon.Col = 2
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("244", CStr(g_rst_Princi!SEGCON_SITUAC))
      
      If g_rst_Princi!SEGCON_SITUAC = 1 Then
         l_int_AprCon = 1
      End If
      
      'Descripci�n Levantamiento Condiciones
      grd_LisCon.Col = 3
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSLEV & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisCon.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisCon)
   Call grd_LisCon_Click
End Sub

Private Sub grd_LisCon_Click()
   Dim r_str_FecOcu     As String
   Dim r_str_HorOcu     As String
   Dim r_str_DesOcu     As String

   If grd_LisCon.Rows > 0 Then
      grd_LisCon.Col = 0
      pnl_InsCon.Caption = grd_LisCon.Text
   
      grd_LisCon.Col = 1
      txt_ObsCon.Text = grd_LisCon.Text
      
      grd_LisCon.Col = 3
      txt_LevCon.Text = grd_LisCon.Text
      
      Call gs_RefrescaGrid(grd_LisCon)
   End If
End Sub

Private Sub grd_LisCon_SelChange()
   If grd_LisCon.Rows > 2 Then
      grd_LisCon.RowSel = grd_LisCon.Row
   End If
   
   Call grd_LisCon_Click
End Sub

Private Sub grd_LisEva_SelChange()
   If grd_LisEva.Rows > 2 Then
      grd_LisEva.RowSel = grd_LisEva.Row
   End If
End Sub

Private Sub grd_LisExc_Click()
   Dim r_str_FecExc     As String
   Dim r_str_HorExc     As String
   Dim r_str_InsExc     As String

   If grd_LisExc.Rows > 0 Then
      grd_LisExc.Col = 0
      r_str_FecExc = grd_LisExc.Text
      
      grd_LisExc.Col = 1
      r_str_HorExc = grd_LisExc.Text
      
      grd_LisExc.Col = 2
      r_str_InsExc = grd_LisExc.Text
      
      pnl_DesExc.Caption = "D�a: " & r_str_FecExc & " - " & r_str_HorExc & " hrs. - " & r_str_InsExc
   
      grd_LisExc.Col = 3
      txt_ObsExc.Text = grd_LisExc.Text
      
      grd_LisExc.Col = 4
      pnl_TipAut.Caption = grd_LisExc.Text
      
      Call gs_RefrescaGrid(grd_LisExc)
   Else
      pnl_DesExc.Caption = ""
      txt_ObsExc.Text = ""
      pnl_TipAut.Caption = ""
   End If
End Sub

Private Sub grd_LisExc_SelChange()
   If grd_LisExc.Rows > 2 Then
      grd_LisExc.RowSel = grd_LisExc.Row
   End If
   
   Call grd_LisExc_Click
End Sub

Private Sub grd_LisOcu_Click()
   Dim r_str_FecOcu     As String
   Dim r_str_HorOcu     As String
   Dim r_str_DesOcu     As String

   If grd_LisOcu.Rows > 0 Then
      grd_LisOcu.Col = 0
      r_str_FecOcu = grd_LisOcu.Text
      
      grd_LisOcu.Col = 1
      r_str_HorOcu = grd_LisOcu.Text
      
      grd_LisOcu.Col = 2
      r_str_DesOcu = grd_LisOcu.Text
      
      pnl_DesOcu.Caption = "D�a: " & r_str_FecOcu & " - " & r_str_HorOcu & " hrs. - " & r_str_DesOcu
   
      grd_LisOcu.Col = 3
      txt_Observ.Text = grd_LisOcu.Text
      
      grd_LisOcu.Col = 4
      txt_Descar.Text = grd_LisOcu.Text
      
      Call gs_RefrescaGrid(grd_LisOcu)
   End If
End Sub

Private Sub grd_LisOcu_SelChange()
   If grd_LisOcu.Rows > 2 Then
      grd_LisOcu.RowSel = grd_LisOcu.Row
   End If
   
   Call grd_LisOcu_Click
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub fs_Buscar_DatEva()
   txt_ObsEva.Text = ""
   Call gs_LimpiaGrid(grd_LisEva)
   
   g_str_Parame = "SELECT * FROM TRA_EVACOF WHERE "
   g_str_Parame = g_str_Parame & "EVACOF_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Fecha Env�o"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECENV))
   
      If g_rst_Princi!EVACOF_FECREC > 0 Then
         If moddat_g_str_CodPrd = "003" Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Nro. Operaci�n Mivivienda"
      
            grd_LisEva.Col = 1
            grd_LisEva.Text = Trim(g_rst_Princi!EVACOF_CODMV1 & "")
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Fecha Aprobaci�n Mivivienda"
      
            grd_LisEva.Col = 1
            grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_APRMVI))
         End If
      
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Nro. Carta COFIDE"
   
         grd_LisEva.Col = 1
         grd_LisEva.Text = Trim(g_rst_Princi!EVACOF_NUMCAR & "")
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Fecha Recepci�n Carta COFIDE"
   
         grd_LisEva.Col = 1
         grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECREC))
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Nro. Operaci�n COFIDE"
   
         grd_LisEva.Col = 1
         grd_LisEva.Text = Trim(g_rst_Princi!EVACOF_CODMVI & "")
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Fecha Desembolso COFIDE"
   
         grd_LisEva.Col = 1
         grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECDES))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Importe Desembolsado"
   
         grd_LisEva.Col = 1
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACOF_MTODES, 12, 2)
         
         txt_ObsEva.Text = Trim(g_rst_Princi!EVACOF_OBSERV & "")
      End If
      
      Call gs_UbiIniGrid(grd_LisEva)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_Descar_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_LevCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsEva_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_DatTas()
   Call gs_LimpiaGrid(grd_Listad(5))
   
   grd_Listad(5).Redraw = False
   
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Empresa Peritaje"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("507", g_rst_Princi!EVATAS_CODEMP)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Nombre Perito"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = Trim(g_rst_Princi!EVATAS_NOMPER & "")
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "C�digo REPEV SBS"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = Trim(g_rst_Princi!EVATAS_CODPER & "")
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Nro. de Informe"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = Trim(g_rst_Princi!EVATAS_NUMINF & "")
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Fecha Evaluaci�n"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "A�o de Construcci�n"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = CStr(g_rst_Princi!EVATAS_ANOCON)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Nro. de Pisos"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = CStr(g_rst_Princi!EVATAS_NUMPIS)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Nro. de S�tanos"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = CStr(g_rst_Princi!EVATAS_NUMSOT)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Tipo de Inmueble"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("221", CStr(g_rst_Princi!EVATAS_TIPINM))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Uso de Inmueble"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("222", CStr(g_rst_Princi!EVATAS_USOINM))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Material de Construcci�n"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("223", CStr(g_rst_Princi!EVATAS_MATCON))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Tipo de Moneda"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!EVATAS_TIPMON))
      
      'Total
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Area Terreno (Total)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM + g_rst_Princi!EVATAS_ARETER_ES1 + g_rst_Princi!EVATAS_ARETER_ES2 + g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Area Construida (Total)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM + g_rst_Princi!EVATAS_ARECON_ES1 + g_rst_Princi!EVATAS_ARECON_ES2 + g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Suma Asegurada (Total)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Valor Comercial (Total)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM + g_rst_Princi!EVATAS_VALCOM_ES1 + g_rst_Princi!EVATAS_VALCOM_ES2 + g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Valor Realizaci�n (Total)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM + g_rst_Princi!EVATAS_VALREA_ES1 + g_rst_Princi!EVATAS_VALREA_ES2 + g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Valor Terreno (Total)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM + g_rst_Princi!EVATAS_VALTER_ES1 + g_rst_Princi!EVATAS_VALTER_ES2 + g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Valor Edificaci�n (Total)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM + g_rst_Princi!EVATAS_VALEDI_ES1 + g_rst_Princi!EVATAS_VALEDI_ES2 + g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Valor Areas Comunes (Total)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM + g_rst_Princi!EVATAS_VALACO_ES1 + g_rst_Princi!EVATAS_VALACO_ES2 + g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
   
      'Inmueble
      grd_Listad(5).Rows = grd_Listad(5).Rows + 2
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).Text = "Area Terreno (Inmueble)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM, 12, 2) & " m2"
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).Text = "Area Construida (Inmueble)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM, 12, 2) & " m2"
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).Text = "Suma Asegurada (Inmueble)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM, 12, 2)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).Text = "Valor Comercial (Inmueble)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM, 12, 2)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).Text = "Valor Realizaci�n (Inmueble)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM, 12, 2)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).Text = "Valor Terreno (Inmueble)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM, 12, 2)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).Text = "Valor Edificaci�n (Inmueble)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).Text = "Valor Areas Comunes (Inmueble)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM, 12, 2)
   
      'Estacionamiento 1
      If g_rst_Princi!EVATAS_FLGEST_ES1 = 1 Then
         grd_Listad(5).Rows = grd_Listad(5).Rows + 2
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Area Terreno (Estac. 1)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES1, 12, 2) & " m2"
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Area Construida (Estac. 1)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES1, 12, 2) & " m2"
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Suma Asegurada (Estac. 1)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES1, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Valor Comercial (Estac. 1)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES1, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Valor Realizaci�n (Estac. 1)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES1, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Valor Terreno (Estac. 1)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES1, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Valor Edificaci�n (Estac. 1)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES1, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Valor Areas Comunes (Estac. 1)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES1, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_ES2 = 1 Then
         grd_Listad(5).Rows = grd_Listad(5).Rows + 2
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).Text = "Area Terreno (Estac. 2)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES2, 12, 2) & " m2"
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).Text = "Area Construida (Estac. 2)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES2, 12, 2) & " m2"
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).Text = "Suma Asegurada (Estac. 2)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES2, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).Text = "Valor Comercial (Estac. 2)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES2, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).Text = "Valor Realizaci�n (Estac. 2)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES2, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).Text = "Valor Terreno (Estac. 2)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES2, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).Text = "Valor Edificaci�n (Estac. 2)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES2, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).Text = "Valor Areas Comunes (Estac. 2)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES2, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_DEP = 1 Then
         grd_Listad(5).Rows = grd_Listad(5).Rows + 2
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Area Terreno (Dep�sito)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Area Construida (Dep�sito)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Suma Asegurada (Dep�sito)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Valor Comercial (Dep�sito)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Valor Realizaci�n (Dep�sito)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Valor Terreno (Dep�sito)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Valor Edificaci�n (Dep�sito)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).Text = "Valor Areas Comunes (Dep�sito)"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   grd_Listad(5).Redraw = True

   Call gs_UbiIniGrid(grd_Listad(5))
End Sub

Private Sub fs_DatSeg()
   Call gs_LimpiaGrid(grd_Listad(6))
   
   grd_Listad(6).Redraw = False
   
   'Seguros
   g_str_Parame = "SELECT * FROM TRA_EVASEG WHERE "
   g_str_Parame = g_str_Parame & "EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Empresa de Seguros"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!EVASEG_ESGDES & "")
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 2
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Tipo de Seguro Desgravamen"

      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!EVASEG_ESGDES, g_rst_Princi!EVASEG_TIPSEG)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Fecha Evaluaci�n (Seg. Desgravamen)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVADES))
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Tipo de Valor (Seg. Desgravamen)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPDES))
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Valor a Aplicar"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = Format(g_rst_Princi!EVASEG_FOIDES, "###,###,##0.000000")
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 2
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Fecha Evaluaci�n (Seg. Inmueble)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAVIV))
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Tipo de Valor (Seg. Inmueble)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPVIV))
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Valor a Aplicar"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = Format(g_rst_Princi!EVASEG_FOIVIV, "###,###,##0.000000")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_Listad(6).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(6))
End Sub

Private Sub fs_DatLeg()
   txt_InfLeg.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad(7))
   Call gs_LimpiaGrid(grd_Listad(8))

   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      txt_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLG1 & "") & Trim(g_rst_Princi!EVALEG_INFLG2 & "") & Trim(g_rst_Princi!EVALEG_INFLG3 & "") & Trim(g_rst_Princi!EVALEG_INFLG4 & "")
      
      If g_rst_Princi!EVALEG_FECCVT > 0 Then
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Fecha Firma Contrato Compra Venta"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCVT))
         
         If Not IsNull(g_rst_Princi!EVALEG_TCASBS) Then
            If g_rst_Princi!EVALEG_TCASBS > 0 Then
               grd_Listad(7).Rows = grd_Listad(7).Rows + 1
               grd_Listad(7).Row = grd_Listad(7).Rows - 1
               grd_Listad(7).Col = 0
               grd_Listad(7).Text = "Tipo de Cambio SBS"
               
               grd_Listad(7).Col = 1
               grd_Listad(7).Text = Format(g_rst_Princi!EVALEG_TCASBS, "###,##0.0000")
            End If
         End If
      
         If g_rst_Princi!EVALEG_TCACVT > 0 Then
            grd_Listad(7).Rows = grd_Listad(7).Rows + 1
            grd_Listad(7).Row = grd_Listad(7).Rows - 1
            grd_Listad(7).Col = 0
            grd_Listad(7).Text = "Tipo de Cambio aplicado"
            
            grd_Listad(7).Col = 1
            grd_Listad(7).Text = Format(g_rst_Princi!EVALEG_TCACVT, "###,##0.0000")
         End If
      End If
      
      If g_rst_Princi!EVALEG_FIRCON > 0 Then
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Fecha Firma Contrato"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
      
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Notaria"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("509", g_rst_Princi!EVALEG_CODNOT)
      
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Representante Legal 1"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG1)
      
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Representante Legal 2"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Monto Hipoteca"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONHIP) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_MTOHIP, 12, 2)
      End If
      
      If grd_Listad(7).Rows > 0 Then
         Call gs_UbiIniGrid(grd_Listad(7))
      End If
      
      If g_rst_Princi!EVALEG_FECBLQ_INM > 0 Then
         grd_Listad(8).Rows = grd_Listad(8).Rows + 1
         grd_Listad(8).Row = grd_Listad(8).Rows - 1
         grd_Listad(8).Col = 0
         grd_Listad(8).Text = "Sede Registral"
         
         grd_Listad(8).Col = 1
         grd_Listad(8).Text = moddat_gf_Consulta_ParDes("511", CStr(g_rst_Princi!EVALEG_SEDREG))
         
         grd_Listad(8).Rows = grd_Listad(8).Rows + 1
         grd_Listad(8).Row = grd_Listad(8).Rows - 1
         grd_Listad(8).Col = 0
         grd_Listad(8).Text = "Fecha Bloqueo (Inmueble)"
         
         grd_Listad(8).Col = 1
         grd_Listad(8).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_INM))
         
         grd_Listad(8).Rows = grd_Listad(8).Rows + 1
         grd_Listad(8).Row = grd_Listad(8).Rows - 1
         grd_Listad(8).Col = 0
         grd_Listad(8).Text = "Doc. Registral (Inmueble)"
         
         grd_Listad(8).Col = 1
         grd_Listad(8).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_INM)
         
         Select Case g_rst_Princi!EVALEG_TIPDOC_INM
            Case 1
               grd_Listad(8).Text = grd_Listad(8).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_INM & "")
               
            Case 2
               grd_Listad(8).Text = grd_Listad(8).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_INM & "")
               
            Case 3
               grd_Listad(8).Text = grd_Listad(8).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_INM & "") & ")"
         End Select
         
         If g_rst_Princi!EVALEG_FLGEST_ES1 = 1 Then
            grd_Listad(8).Rows = grd_Listad(8).Rows + 2
            grd_Listad(8).Row = grd_Listad(8).Rows - 1
            grd_Listad(8).Col = 0
            grd_Listad(8).Text = "Fecha Bloqueo (Estac. 1)"
            
            grd_Listad(8).Col = 1
            grd_Listad(8).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES1))
            
            grd_Listad(8).Rows = grd_Listad(8).Rows + 1
            grd_Listad(8).Row = grd_Listad(8).Rows - 1
            grd_Listad(8).Col = 0
            grd_Listad(8).Text = "Doc. Registral (Estac. 1)"
            
            grd_Listad(8).Col = 1
            grd_Listad(8).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES1)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES1
               Case 1
                  grd_Listad(8).Text = grd_Listad(8).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES1 & "")
               
               Case 2
                  grd_Listad(8).Text = grd_Listad(8).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES1 & "")
               
               Case 3
                  grd_Listad(8).Text = grd_Listad(8).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES1 & "") & ")"
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_ES2 = 1 Then
            grd_Listad(8).Rows = grd_Listad(8).Rows + 2
            grd_Listad(8).Row = grd_Listad(8).Rows - 1
            grd_Listad(8).Col = 0
            grd_Listad(8).Text = "Fecha Bloqueo (Estac. 2)"
            
            grd_Listad(8).Col = 1
            grd_Listad(8).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES2))
            
            grd_Listad(8).Rows = grd_Listad(8).Rows + 1
            grd_Listad(8).Row = grd_Listad(8).Rows - 1
            grd_Listad(8).Col = 0
            grd_Listad(8).Text = "Doc. Registral (Estac. 2)"
            
            grd_Listad(8).Col = 1
            grd_Listad(8).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES2)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES2
               Case 1
                  grd_Listad(8).Text = grd_Listad(8).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES2 & "")
                  
               Case 2
                  grd_Listad(8).Text = grd_Listad(8).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES2 & "")
                  
               Case 3
                  grd_Listad(8).Text = grd_Listad(8).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES2 & "") & ")"
                  
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_DEP = 1 Then
            grd_Listad(8).Rows = grd_Listad(8).Rows + 2
            grd_Listad(8).Row = grd_Listad(8).Rows - 1
            grd_Listad(8).Col = 0
            grd_Listad(8).Text = "Fecha Bloqueo (Dep�sito)"
            
            grd_Listad(8).Col = 1
            grd_Listad(8).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_DEP))
            
            grd_Listad(8).Rows = grd_Listad(8).Rows + 1
            grd_Listad(8).Row = grd_Listad(8).Rows - 1
            grd_Listad(8).Col = 0
            grd_Listad(8).Text = "Doc. Registral (Dep�sito)"
            
            grd_Listad(8).Col = 1
            grd_Listad(8).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_DEP)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_DEP
               Case 1
                  grd_Listad(8).Text = grd_Listad(8).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_DEP & "")
               
               Case 2
                  grd_Listad(8).Text = grd_Listad(8).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_DEP & "")
                  
               Case 3
                  grd_Listad(8).Text = grd_Listad(8).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_DEP & "") & ")"
            End Select
         End If
      End If
      
      If grd_Listad(8).Rows > 0 Then
         Call gs_UbiIniGrid(grd_Listad(8))
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_EvaCre()
   Call gs_LimpiaGrid(grd_Listad(4))
   
   'Obteniendo Ingreso Neto
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
   grd_Listad(4).Text = "Total Ingreso L�quido Neto S/."
   
   grd_Listad(4).Col = 1
   grd_Listad(4).CellFontName = "Lucida Console"
   grd_Listad(4).CellFontSize = 8
   grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
   grd_Listad(4).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGNET, 12, 2)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'Obteniendo Cuota Aceptada
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Cuota (S/.)"

   grd_Listad(4).Col = 1
   grd_Listad(4).CellFontName = "Lucida Console"
   grd_Listad(4).CellFontSize = 8
   grd_Listad(4).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_SOL, 12, 2)

   grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   grd_Listad(4).Row = grd_Listad(4).Rows - 1
   grd_Listad(4).Col = 0
   grd_Listad(4).Text = "Cuota (Moneda Prest.)"

   grd_Listad(4).Col = 1
   grd_Listad(4).CellFontName = "Lucida Console"
   grd_Listad(4).CellFontSize = 8
   grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_MPR, 12, 2)

   If g_rst_Princi!SOLMAE_TIPMON <> 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).Text = "Tipo de Cambio"
   
      grd_Listad(4).Col = 1
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_TCAMPR_APR, 14, 4)
   End If

   Call gs_UbiIniGrid(grd_Listad(4))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_ObsExc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_Envia_CorEle(ByVal p_Asunto As String, ByVal p_Mensaje As String)
   Dim r_str_Cadena     As String
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Usuario de Seguimiento
   r_str_Cadena = moddat_gf_Buscar_DirEle_Codigo(moddat_g_str_CodEjeSeg)
   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(moddat_g_str_CodEjeSeg)
   
   'Consejero Hipotecario
   r_str_Cadena = moddat_gf_Buscar_DirEle_Codigo(moddat_g_str_CodConHip)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Seguimiento
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(130)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Ventas
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(120)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director Comercial
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(100)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Operaciones
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(220)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director de Producci�n
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(200)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director de Administracion
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(300)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director General
   'r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(10)
   'If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
   '   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   '   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   'End If
   
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
   'Call modsec_gs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, "", "")
End Sub




