VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_Tra_EvaSeg_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9630
   ClientLeft      =   660
   ClientTop       =   1635
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_281.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   17013
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
         Top             =   4080
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
            Tab             =   2
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento de Instancias"
            TabPicture(0)   =   "OpeTra_frm_281.frx":000C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Label7"
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(2)=   "Label11"
            Tab(0).Control(3)=   "pnl_DesOcu"
            Tab(0).Control(4)=   "SSPanel5"
            Tab(0).Control(5)=   "SSPanel14"
            Tab(0).Control(6)=   "SSPanel13"
            Tab(0).Control(7)=   "grd_LisOcu"
            Tab(0).Control(8)=   "SSPanel10"
            Tab(0).Control(9)=   "txt_Observ"
            Tab(0).Control(10)=   "txt_Descar"
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "OpeTra_frm_281.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label6"
            Tab(1).Control(1)=   "Label3"
            Tab(1).Control(2)=   "Label4"
            Tab(1).Control(3)=   "lbl_motivo"
            Tab(1).Control(4)=   "pnl_motivo"
            Tab(1).Control(5)=   "pnl_TipAut"
            Tab(1).Control(6)=   "pnl_DesExc"
            Tab(1).Control(7)=   "SSPanel16"
            Tab(1).Control(8)=   "SSPanel15"
            Tab(1).Control(9)=   "SSPanel12"
            Tab(1).Control(10)=   "SSPanel11"
            Tab(1).Control(11)=   "SSPanel9"
            Tab(1).Control(12)=   "grd_LisExc"
            Tab(1).Control(13)=   "txt_ObsExc"
            Tab(1).ControlCount=   14
            TabCaption(2)   =   "Aprobaci�n Condicionada"
            TabPicture(2)   =   "OpeTra_frm_281.frx":0044
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Label15"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "Label14"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "Label12"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "pnl_InsCon"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "SSPanel20"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).Control(5)=   "SSPanel19"
            Tab(2).Control(5).Enabled=   0   'False
            Tab(2).Control(6)=   "SSPanel18"
            Tab(2).Control(6).Enabled=   0   'False
            Tab(2).Control(7)=   "grd_LisCon"
            Tab(2).Control(7).Enabled=   0   'False
            Tab(2).Control(8)=   "SSPanel17"
            Tab(2).Control(8).Enabled=   0   'False
            Tab(2).Control(9)=   "txt_ObsCon"
            Tab(2).Control(9).Enabled=   0   'False
            Tab(2).Control(10)=   "txt_LevCon"
            Tab(2).Control(10).Enabled=   0   'False
            Tab(2).ControlCount=   11
            Begin VB.TextBox txt_Descar 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Text            =   "OpeTra_frm_281.frx":0060
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_Observ 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Text            =   "OpeTra_frm_281.frx":0064
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   975
               Left            =   -73770
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Text            =   "OpeTra_frm_281.frx":0068
               Top             =   1980
               Width           =   10095
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Text            =   "OpeTra_frm_281.frx":006C
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsCon 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "OpeTra_frm_281.frx":0070
               Top             =   1980
               Width           =   10005
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   45
               Left            =   -74970
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
               Left            =   -74970
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
               Left            =   -74940
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
               Left            =   -72600
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
               Left            =   -73770
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
               Left            =   -73680
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
               Cols            =   6
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
               Left            =   -73770
               TabIndex        =   20
               Top             =   1650
               Width           =   10095
               _Version        =   65536
               _ExtentX        =   17806
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
               Left            =   -73770
               TabIndex        =   21
               Top             =   2970
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "INGRESO A INSTANCIA"
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
               Left            =   30
               TabIndex        =   22
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisCon 
               Height          =   855
               Left            =   30
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
               Left            =   60
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
               Left            =   9390
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
               Left            =   2790
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
               Left            =   1320
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
                  Size            =   8.26
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
            Begin Threed.SSPanel pnl_motivo 
               Height          =   315
               Left            =   -69150
               TabIndex        =   70
               Top             =   2970
               Width           =   5475
               _Version        =   65536
               _ExtentX        =   9657
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "MOTIVO DE EXCEPCION"
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
               Begin Threed.SSPanel SSPanel21 
                  Height          =   315
                  Left            =   6090
                  TabIndex        =   71
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   5205
                  _Version        =   65536
                  _ExtentX        =   9181
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "INGRESOS 4A CATEG. NO SE PUEDEN CONFIRMAR "
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
            End
            Begin VB.Label lbl_motivo 
               Caption         =   "Motivo:"
               Height          =   255
               Left            =   -69840
               TabIndex        =   72
               Top             =   3030
               Width           =   645
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   36
               Top             =   2640
               Width           =   1035
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   35
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observaci�n:"
               Height          =   495
               Left            =   -74940
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
               Left            =   -74940
               TabIndex        =   31
               Top             =   2970
               Width           =   1095
            End
            Begin VB.Label Label12 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   615
               Left            =   60
               TabIndex        =   30
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   60
               TabIndex        =   29
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobaci�n:"
               Height          =   495
               Left            =   60
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
            BackColor       =   192
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   690
            TabIndex        =   59
            Top             =   30
            Width           =   6195
            _Version        =   65536
            _ExtentX        =   10927
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Solicitud de Cr�dito Hipotecario"
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
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   690
            TabIndex        =   60
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluaci�n de Seguros"
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_281.frx":0074
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel36 
         Height          =   645
         Left            =   0
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
         Begin VB.CommandButton cmd_EnvMail 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_281.frx":037E
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Envio de Correo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_281.frx":0688
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   585
            Left            =   3030
            Picture         =   "OpeTra_frm_281.frx":0ACA
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   585
            Left            =   3630
            Picture         =   "OpeTra_frm_281.frx":0DD4
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_AprCon 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_281.frx":1216
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Aprobaci�n con Condici�n"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Excepc 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_281.frx":1520
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Registro de Excepci�n"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueObs 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_281.frx":182A
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Registro de Observaci�n"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Evalua 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_281.frx":1C6C
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Registro de Tasas"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   48
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
            TabIndex        =   49
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   50
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
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
         Begin Threed.SSPanel pnl_FecSol 
            Height          =   315
            Left            =   9450
            TabIndex        =   51
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
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   54
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   53
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   8070
            TabIndex        =   52
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1995
         Left            =   30
         TabIndex        =   55
         Top             =   7590
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3519
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
            Height          =   675
            Left            =   1350
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   56
            Text            =   "OpeTra_frm_281.frx":1F76
            Top             =   1260
            Width           =   10095
         End
         Begin MSFlexGridLib.MSFlexGrid grd_LisEva 
            Height          =   1155
            Left            =   30
            TabIndex        =   57
            Top             =   60
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   2037
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label5 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   58
            Top             =   1260
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1785
         Left            =   30
         TabIndex        =   61
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3149
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
            Height          =   1665
            Left            =   60
            TabIndex        =   62
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   2937
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            Tab             =   4
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "OpeTra_frm_281.frx":1F7A
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "C�nyuge"
            TabPicture(1)   =   "OpeTra_frm_281.frx":1F96
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Apoderado"
            TabPicture(2)   =   "OpeTra_frm_281.frx":1FB2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Inmueble"
            TabPicture(3)   =   "OpeTra_frm_281.frx":1FCE
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos Cr�dito"
            TabPicture(4)   =   "OpeTra_frm_281.frx":1FEA
            Tab(4).ControlEnabled=   -1  'True
            Tab(4).Control(0)=   "grd_Listad(4)"
            Tab(4).Control(0).Enabled=   0   'False
            Tab(4).Control(1)=   "txt_ObsSol"
            Tab(4).Control(1).Enabled=   0   'False
            Tab(4).ControlCount=   2
            Begin VB.TextBox txt_ObsSol 
               Height          =   405
               Left            =   1290
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   63
               Text            =   "OpeTra_frm_281.frx":2006
               Top             =   1200
               Width           =   10005
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1245
               Index           =   0
               Left            =   -74940
               TabIndex        =   64
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
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
               Height          =   1245
               Index           =   1
               Left            =   -74940
               TabIndex        =   65
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
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
               Height          =   1245
               Index           =   2
               Left            =   -74940
               TabIndex        =   66
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
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
               Height          =   1245
               Index           =   3
               Left            =   -74940
               TabIndex        =   67
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
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
               Height          =   825
               Index           =   4
               Left            =   60
               TabIndex        =   68
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   1455
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label9 
               Caption         =   "Observaciones:"
               Height          =   300
               Left            =   -74910
               TabIndex        =   69
               Top             =   1200
               Width           =   1155
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_EvaSeg_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_AprCon     As Integer
Dim l_arr_Mtz()      As moddat_g_tpo_DatCom

Private Sub cmd_AprCon_Click()
   Dim r_int_TipDoc     As Integer
   Dim r_int_CodAct     As Integer
   Dim r_int_Contad     As Integer
   Dim r_int_FlgDoc     As Integer
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodGrp     As String
   Dim r_str_CodIte     As String

   If grd_LisEva.Rows = 0 Then
      MsgBox "No se ha registrado informaci�n de las Tasas de Seguro.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If moddat_gf_Valida_Observ(moddat_g_str_NumSol, 42) Then
      MsgBox "La solicitud presenta Observaciones pendientes de descargo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If pnl_AprCon.Visible Then
      MsgBox "La solicitud presenta Aprobaci�n Condicionada pendiente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("�Est� seguro de aprobar esta instancia de Evaluaci�n?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_str_Observ = ""
   moddat_g_int_FlgAct_1 = 1
   
   frm_RecSol_56.Show 1

   If moddat_g_int_FlgAct_1 = 1 Then
      Exit Sub
   End If
   
   'Creando Aprobaci�n Condicionada
   If Not moddat_gf_Inserta_AprCon(moddat_g_str_NumSol, 42, moddat_g_str_Observ) Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 42)))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 42, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 42, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Verificar si aprobo Tasaci�n
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 41"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   g_rst_Genera.MoveFirst

   If g_rst_Genera!SEGUIM_SITUAC <> 1 Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluaci�n.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
   Else
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      'Inserta Nueva Instancia de Evaluaci�n
      If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 51) Then
         Exit Sub
      End If
         
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 51, 11, 0, "", 0, 0) Then
         Exit Sub
      End If
      
      'Actualizando en Tabla de Cr�ditos
      If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, 51) Then
         Exit Sub
      End If
      
      'Enviando Correo Electr�nico
      modgen_g_str_Mail_Asunto = "TASACION DEL INMUEBLE Y EVALUACION DE SEGUROS - APROBACION CONDICIONADA (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, True, False)
      
      MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluaci�n.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
   End If
   
   Unload Me
End Sub

Private Sub cmd_Aprueb_Click()
   Dim r_int_TipDoc     As Integer
   Dim r_int_CodAct     As Integer
   Dim r_int_Contad     As Integer
   Dim r_int_FlgDoc     As Integer
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodGrp     As String
   Dim r_str_CodIte     As String

   If grd_LisEva.Rows = 0 Then
      MsgBox "No se ha registrado informaci�n de las Tasas de Seguro.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If moddat_gf_Valida_Observ(moddat_g_str_NumSol, 42) Then
      MsgBox "La solicitud presenta Observaciones pendientes de descargo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If pnl_AprCon.Visible Then
      MsgBox "La solicitud presenta Aprobaci�n Condicionada pendiente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("�Est� seguro de aprobar esta instancia de Evaluaci�n?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 42)))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 42, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 42, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Verificar si aprobo Tasaci�n
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 41"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   g_rst_Genera.MoveFirst

   If g_rst_Genera!SEGUIM_SITUAC <> 1 Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluaci�n.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
   Else
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      'Inserta Nueva Instancia de Evaluaci�n
      If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 51) Then
         Exit Sub
      End If
         
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 51, 11, 0, "", 0, 0) Then
         Exit Sub
      End If
      
      'Actualizando en Tabla de Cr�ditos
      If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, 51) Then
         Exit Sub
      End If
      
      'Enviando Correo Electr�nico
      modgen_g_str_Mail_Asunto = "TASACION DEL INMUEBLE Y EVALUACION DE SEGUROS - APROBACION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, True, False)
      
      MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluaci�n.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
   End If
   
   Unload Me
End Sub

Private Sub cmd_EnvMail_Click()
   frm_Tra_EvaSeg_04.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
   End If
End Sub

Private Sub cmd_Evalua_Click()
   moddat_g_int_FlgAct = 1
   frm_Tra_EvaSeg_03.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      
      'Call fs_DatCre
      'Buscar Informaci�n del Cr�dito
      Call modmip_gs_DatCre(grd_Listad(4), l_arr_Mtz)
      txt_ObsSol.Text = l_arr_Mtz(0).DatCom_Observ
      moddat_g_str_CodEjeSeg = l_arr_Mtz(0).DatCom_EjeSeg
      moddat_g_str_CodConHip = l_arr_Mtz(0).DatCom_ConHip
      moddat_g_str_FecIng = l_arr_Mtz(0).DatCom_FecSol
   
      Call fs_Buscar_DatEva      'Buscando Informaci�n de Evaluaci�n ya registrada
      Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
      
      'Buscando Excepciones
      Call fs_Buscar_LisExc
      
      'Si no hay Excepciones aplicadas
      If grd_LisExc.Rows = 0 Then
         tab_Seguim.TabVisible(1) = False
      End If
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Excepc_Click()
Dim r_int_NumExc     As Integer

   moddat_g_str_Observ = ""
   moddat_g_int_TipAut = 0
   moddat_g_int_FlgAct_1 = 1
   
   frm_RecSol_55.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
   
      'Generando N�mero de Excepci�n
      r_int_NumExc = 0
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT COUNT(SEGEXC_NUMSOL) AS NUMREG FROM TRA_SEGEXC "
      g_str_Parame = g_str_Parame & " WHERE SEGEXC_NUMSOL = '" & moddat_g_str_NumSol & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         r_int_NumExc = g_rst_Princi!NUMREG
      End If
         
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
         
      r_int_NumExc = r_int_NumExc + 1
      
      'Grabando en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 42, 18, 0, "", 0, 0) Then
         Exit Sub
      End If
      
      'Grabando en Detalle de Excepciones
      If Not moddat_gf_Inserta_SegExc(moddat_g_str_NumSol, 42, r_int_NumExc, moddat_g_str_Observ, moddat_g_int_TipAut) Then
         Exit Sub
      End If
      
      'Enviando Correo Electr�nico
      modgen_g_str_Mail_Asunto = "EVALUACION DE SEGUROS - EXCEPCION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False)
      
      Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
      Call fs_Buscar_LisExc      'Buscando Excepciones
   
      'Si no hay Excepciones aplicadas
      If grd_LisExc.Rows = 0 Then
         tab_Seguim.TabVisible(1) = False
      Else
         tab_Seguim.TabVisible(1) = True
      End If
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_NueObs_Click()
   Dim r_int_NumObs     As Integer
   
   moddat_g_int_TipObs = 0
   moddat_g_str_Observ = ""
   moddat_g_int_FlgAct_1 = 1
   
   frm_RecSol_54.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      
      If moddat_g_int_TipObs = 1 Then
         'Generando N�mero de Observaci�n
         r_int_NumObs = 0
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT * FROM TRA_SEGDET "
         g_str_Parame = g_str_Parame & " WHERE SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' "
         g_str_Parame = g_str_Parame & "   AND SEGDET_CODINS = 42 "
         g_str_Parame = g_str_Parame & "   AND SEGDET_CODOCU = 21 "
         g_str_Parame = g_str_Parame & " ORDER BY SEGDET_NUMOBS DESC"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
             Exit Sub
         End If
      
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            g_rst_Princi.MoveFirst
            
            Do While Not g_rst_Princi.EOF
               r_int_NumObs = r_int_NumObs + 1
               g_rst_Princi.MoveNext
            Loop
         End If
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         
         r_int_NumObs = r_int_NumObs + 1
   
         'Grabando en Detalle de Seguimiento
         If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 42, 21, CStr(r_int_NumObs), moddat_g_str_Observ, 1, 0) Then
            Exit Sub
         End If
         
         'Actualizando en Instancia si es una Observaci�n
         If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 42, 0, 3, 2) Then
            Exit Sub
         End If
         
         modgen_g_str_Mail_Asunto = "EVALUACION DE SEGUROS - OBSERVACION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      Else
         'Grabando en Detalle de Seguimiento
         If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 42, 17, 0, moddat_g_str_Observ, 0, 0) Then
            Exit Sub
         End If
         
         modgen_g_str_Mail_Asunto = "EVALUACION DE SEGUROS - COMENTARIO (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      End If
      
      'Enviando Correo Electr�nico
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False)
      
      'Cargando Datos de Seguimiento
      Call fs_Buscar_LisOcu
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Rechaz_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = 42
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 42)))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 42, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 42, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      'Actualizando Rechazo en Tabla de Cr�ditos
      If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      modgen_g_str_Mail_Asunto = "EVALUACION DE SEGUROS - RECHAZO (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_gf_Consulta_ParDes("003", moddat_g_int_MotRec) & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False)
   
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluaci�n.", vbInformation, modgen_g_str_NomPlt
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
   moddat_g_int_CodIns = 42
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Informaci�n del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Informaci�n del C�nyuge
   Call modmip_gs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(2))         'Buscar Informaci�n del Apoderado
   Call modmip_gs_DatInm(grd_Listad(3), False)                                            'Buscar Informaci�n del Inmueble
      
   'Buscar Informaci�n del Cr�dito
   Call modmip_gs_DatCre(grd_Listad(4), l_arr_Mtz)
   txt_ObsSol.Text = l_arr_Mtz(0).DatCom_Observ
   moddat_g_str_CodEjeSeg = l_arr_Mtz(0).DatCom_EjeSeg
   moddat_g_str_CodConHip = l_arr_Mtz(0).DatCom_ConHip
   moddat_g_str_FecIng = l_arr_Mtz(0).DatCom_FecSol
   
   'Call fs_DatCre
   Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
   Call fs_Buscar_LisCon      'Buscando Aprobaciones Condicionadas
   
   'Buscando Excepciones
   Call fs_Buscar_LisExc
   
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
   For r_int_Contad = 0 To 4
      grd_Listad(r_int_Contad).ColWidth(0) = 3000
      grd_Listad(r_int_Contad).ColWidth(1) = 7940
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
   grd_LisExc.ColWidth(5) = 0
   grd_LisExc.ColAlignment(0) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(1) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(2) = flexAlignLeftCenter
   grd_LisExc.ColAlignment(3) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_LisExc)

   pnl_DesExc.Caption = ""
   txt_ObsExc.Text = ""
   pnl_TipAut.Caption = ""
   pnl_motivo.Caption = ""

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

Private Sub fs_Buscar_LisOcu()
Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisOcu)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_SEGDET "
   g_str_Parame = g_str_Parame & " WHERE SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "   AND SEGDET_CODINS = 42 "
   g_str_Parame = g_str_Parame & " ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
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
   
   g_str_Parame = modgen_gf_Buscar_Excepc
   
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
      
      'Motivo de Excepci�n
      grd_LisExc.Col = 5
      grd_LisExc.Text = Trim(g_rst_Princi!PARDES_DESCRI)
      
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
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_SEGCON "
   g_str_Parame = g_str_Parame & " WHERE SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & " ORDER BY SEGCON_SITUAC ASC, SEGCON_CODINS DESC"
   
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
      
      grd_LisExc.Col = 5
      If LCase(Trim(r_str_InsExc)) = LCase("EVALUACION CREDITICIA") Then
         pnl_motivo.Visible = True
         lbl_motivo.Visible = True
         pnl_motivo.Caption = IIf(grd_LisExc.Text = "0", " ", grd_LisExc.Text)
      Else
         pnl_motivo.Visible = False
         lbl_motivo.Visible = False
         pnl_motivo.Caption = ""
      End If
      
      Call gs_SetFocus(grd_LisExc)
      Call gs_RefrescaGrid(grd_LisExc)
   Else
      pnl_DesExc.Caption = ""
      txt_ObsExc.Text = ""
      pnl_TipAut.Caption = ""
      pnl_motivo.Caption = ""
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
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_EVASEG "
   g_str_Parame = g_str_Parame & " WHERE EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Empresa de Seguros"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!EVASEG_ESGDES & "")
   
      grd_LisEva.Rows = grd_LisEva.Rows + 2
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Tipo de Seguro Desgravamen"

      grd_LisEva.Col = 1
      grd_LisEva.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!EVASEG_ESGDES, g_rst_Princi!EVASEG_TIPSEG)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Fecha Evaluaci�n (Seg. Desgravamen)"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVADES))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Tipo de Valor (Seg. Desgravamen)"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPDES))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Valor a Aplicar"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = Format(g_rst_Princi!EVASEG_FOIDES, "###,###,##0.000000")
      
      grd_LisEva.Rows = grd_LisEva.Rows + 2
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Fecha Evaluaci�n (Seg. Inmueble)"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAVIV))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Tipo de Valor (Seg. Inmueble)"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPVIV))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Valor a Aplicar"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = Format(g_rst_Princi!EVASEG_FOIVIV, "###,###,##0.000000")
      
      txt_ObsEva.Text = Trim(g_rst_Princi!EVASEG_OBSERV & "")
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

Private Sub txt_ObsExc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsRec_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsSol_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsEva_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

