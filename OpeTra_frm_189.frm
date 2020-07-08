VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_Seg_SolHip_65 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10320
   ClientLeft      =   585
   ClientTop       =   1980
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_189.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10335
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   18230
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
         Height          =   3015
         Left            =   30
         TabIndex        =   6
         Top             =   3600
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5318
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
            Height          =   2895
            Left            =   60
            TabIndex        =   7
            Top             =   30
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   5106
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento en Instancia"
            TabPicture(0)   =   "OpeTra_frm_189.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label11"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label7"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "pnl_DesOcu"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel8"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel14"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel13"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "grd_LisOcu"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "txt_Descar"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txt_Observ"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).ControlCount=   10
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "OpeTra_frm_189.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txt_ObsExc"
            Tab(1).Control(1)=   "grd_LisExc"
            Tab(1).Control(2)=   "SSPanel4"
            Tab(1).Control(3)=   "SSPanel5"
            Tab(1).Control(4)=   "SSPanel9"
            Tab(1).Control(5)=   "SSPanel11"
            Tab(1).Control(6)=   "pnl_DesExc"
            Tab(1).Control(7)=   "pnl_TipAut"
            Tab(1).Control(8)=   "pnl_motivo"
            Tab(1).Control(9)=   "lbl_motivo"
            Tab(1).Control(10)=   "Label2"
            Tab(1).Control(11)=   "Label3"
            Tab(1).Control(12)=   "Label4"
            Tab(1).ControlCount=   13
            TabCaption(2)   =   "Aprobación Condicionada"
            TabPicture(2)   =   "OpeTra_frm_189.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txt_ObsCon"
            Tab(2).Control(1)=   "txt_LevCon"
            Tab(2).Control(2)=   "grd_LisCon"
            Tab(2).Control(3)=   "SSPanel18"
            Tab(2).Control(4)=   "SSPanel19"
            Tab(2).Control(5)=   "SSPanel20"
            Tab(2).Control(6)=   "pnl_InsCon"
            Tab(2).Control(7)=   "Label15"
            Tab(2).Control(8)=   "Label14"
            Tab(2).Control(9)=   "Label12"
            Tab(2).ControlCount=   10
            Begin VB.TextBox txt_Observ 
               Height          =   495
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Text            =   "OpeTra_frm_189.frx":0060
               Top             =   1830
               Width           =   10005
            End
            Begin VB.TextBox txt_Descar 
               Height          =   495
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   11
               Text            =   "OpeTra_frm_189.frx":0066
               Top             =   2340
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   675
               Left            =   -73770
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Text            =   "OpeTra_frm_189.frx":006A
               Top             =   1830
               Width           =   10065
            End
            Begin VB.TextBox txt_ObsCon 
               Height          =   495
               Left            =   -73710
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   9
               Text            =   "OpeTra_frm_189.frx":0070
               Top             =   1830
               Width           =   10035
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   465
               Left            =   -73710
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Text            =   "OpeTra_frm_189.frx":0076
               Top             =   2310
               Width           =   10035
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisOcu 
               Height          =   825
               Left            =   30
               TabIndex        =   13
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1455
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
               TabIndex        =   14
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
               TabIndex        =   15
               Top             =   360
               Width           =   8595
               _Version        =   65536
               _ExtentX        =   15161
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Ocurrencia"
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
               Left            =   1230
               TabIndex        =   16
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
               TabIndex        =   17
               Top             =   1500
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
               Height          =   825
               Left            =   -74970
               TabIndex        =   18
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1455
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
            Begin Threed.SSPanel SSPanel4 
               Height          =   285
               Left            =   -74940
               TabIndex        =   19
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Excepción"
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
               Left            =   -69330
               TabIndex        =   20
               Top             =   360
               Width           =   5325
               _Version        =   65536
               _ExtentX        =   9393
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Excepción"
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
               Left            =   -73770
               TabIndex        =   21
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Excepción"
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
               Left            =   -72600
               TabIndex        =   22
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
            Begin Threed.SSPanel pnl_DesExc 
               Height          =   315
               Left            =   -73770
               TabIndex        =   23
               Top             =   1500
               Width           =   10065
               _Version        =   65536
               _ExtentX        =   17754
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
               TabIndex        =   24
               Top             =   2520
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7064
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisCon 
               Height          =   825
               Left            =   -74970
               TabIndex        =   25
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1455
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
               TabIndex        =   26
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
               TabIndex        =   27
               Top             =   360
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Situación"
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
               TabIndex        =   28
               Top             =   360
               Width           =   6615
               _Version        =   65536
               _ExtentX        =   11668
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Condiciones de Aprobación"
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
               TabIndex        =   29
               Top             =   1500
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel pnl_motivo 
               Height          =   315
               Left            =   -68970
               TabIndex        =   59
               Top             =   2520
               Width           =   5265
               _Version        =   65536
               _ExtentX        =   9287
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
                  TabIndex        =   60
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
               Left            =   -69660
               TabIndex        =   61
               Top             =   2580
               Width           =   645
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observación:"
               Height          =   495
               Left            =   60
               TabIndex        =   38
               Top             =   1830
               Width           =   1155
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   60
               TabIndex        =   37
               Top             =   1500
               Width           =   1155
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   60
               TabIndex        =   36
               Top             =   2370
               Width           =   1035
            End
            Begin VB.Label Label2 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   35
               Top             =   2550
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Excepción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   34
               Top             =   1500
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripción:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   33
               Top             =   1830
               Width           =   1155
            End
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobación:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   32
               Top             =   1830
               Width           =   1155
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   31
               Top             =   1500
               Width           =   1155
            End
            Begin VB.Label Label12 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   435
               Left            =   -74940
               TabIndex        =   30
               Top             =   2370
               Width           =   1215
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1755
         Left            =   30
         TabIndex        =   39
         Top             =   1830
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3096
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
            TabIndex        =   40
            Top             =   30
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   2937
            _Version        =   393216
            Style           =   1
            Tabs            =   6
            Tab             =   4
            TabsPerRow      =   6
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "OpeTra_frm_189.frx":007A
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "OpeTra_frm_189.frx":0096
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Apoderado"
            TabPicture(2)   =   "OpeTra_frm_189.frx":00B2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Inmueble"
            TabPicture(3)   =   "OpeTra_frm_189.frx":00CE
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos Crédito"
            TabPicture(4)   =   "OpeTra_frm_189.frx":00EA
            Tab(4).ControlEnabled=   -1  'True
            Tab(4).Control(0)=   "Label5"
            Tab(4).Control(0).Enabled=   0   'False
            Tab(4).Control(1)=   "grd_Listad(4)"
            Tab(4).Control(1).Enabled=   0   'False
            Tab(4).Control(2)=   "txt_ObsSol"
            Tab(4).Control(2).Enabled=   0   'False
            Tab(4).ControlCount=   3
            TabCaption(5)   =   "Ev. Crediticia"
            TabPicture(5)   =   "OpeTra_frm_189.frx":0106
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(5)"
            Tab(5).ControlCount=   1
            Begin VB.TextBox txt_ObsSol 
               Height          =   405
               Left            =   1290
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   53
               Top             =   1200
               Width           =   10005
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1245
               Index           =   0
               Left            =   -74940
               TabIndex        =   41
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
               TabIndex        =   42
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
               TabIndex        =   54
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
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1245
               Index           =   5
               Left            =   -74940
               TabIndex        =   56
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
               TabIndex        =   57
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
               TabIndex        =   58
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
            Begin VB.Label Label5 
               Caption         =   "Observaciones:"
               Height          =   300
               Left            =   60
               TabIndex        =   55
               Top             =   1200
               Width           =   1155
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   43
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
         Begin Threed.SSPanel pnl_AprCon 
            Height          =   555
            Left            =   8520
            TabIndex        =   44
            Top             =   60
            Width           =   2955
            _Version        =   65536
            _ExtentX        =   5212
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   630
            TabIndex        =   45
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud de Crédito Hipotecario"
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
            TabIndex        =   46
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Aceptación del Cliente"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   7920
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
            Left            =   7350
            Top             =   30
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
            Picture         =   "OpeTra_frm_189.frx":0122
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   435
         Left            =   30
         TabIndex        =   47
         Top             =   1380
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   48
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   4440
            TabIndex        =   49
            Top             =   60
            Width           =   7035
            _Version        =   65536
            _ExtentX        =   12409
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
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   225
            Left            =   60
            TabIndex        =   51
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   3690
            TabIndex        =   50
            Top             =   120
            Width           =   645
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   52
         Top             =   720
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_189.frx":042C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CalCuo 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_189.frx":086E
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Calcular Cuota"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_189.frx":0B80
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_189.frx":0FC2
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_189.frx":12CC
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Modificar Dirección del Inmueble"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   3675
         Left            =   30
         TabIndex        =   62
         Top             =   6630
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6482
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
         Begin TabDlg.SSTab SSTab2 
            Height          =   3555
            Left            =   60
            TabIndex        =   63
            Top             =   60
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   6271
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Parámetros el Crédito"
            TabPicture(0)   =   "OpeTra_frm_189.frx":1B96
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel16"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Enlaces"
            TabPicture(1)   =   "OpeTra_frm_189.frx":1BB2
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel12"
            Tab(1).ControlCount=   1
            Begin Threed.SSPanel SSPanel16 
               Height          =   3135
               Left            =   60
               TabIndex        =   64
               Top             =   360
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   5530
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
               Begin VB.ComboBox cmb_TasEsp 
                  Height          =   315
                  Left            =   8910
                  Style           =   2  'Dropdown List
                  TabIndex        =   69
                  Top             =   2070
                  Width           =   2265
               End
               Begin VB.ComboBox cmb_CuoDbl 
                  Height          =   315
                  Left            =   8910
                  Style           =   2  'Dropdown List
                  TabIndex        =   68
                  Top             =   90
                  Width           =   2265
               End
               Begin VB.ComboBox cmb_EmpSeg 
                  Height          =   315
                  Left            =   8910
                  Style           =   2  'Dropdown List
                  TabIndex        =   67
                  Top             =   1410
                  Width           =   2265
               End
               Begin VB.ComboBox cmb_SegDes 
                  Height          =   315
                  Left            =   8910
                  Style           =   2  'Dropdown List
                  TabIndex        =   66
                  Top             =   1740
                  Width           =   2265
               End
               Begin VB.ComboBox cmb_DiaPag 
                  Height          =   315
                  Left            =   8910
                  Style           =   2  'Dropdown List
                  TabIndex        =   65
                  Top             =   1080
                  Width           =   2265
               End
               Begin EditLib.fpLongInteger ipp_PlaAno 
                  Height          =   315
                  Left            =   8910
                  TabIndex        =   70
                  Top             =   420
                  Width           =   2265
                  _Version        =   196608
                  _ExtentX        =   3995
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
                  MaxValue        =   "70"
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
               Begin EditLib.fpLongInteger ipp_PerGra 
                  Height          =   315
                  Left            =   8910
                  TabIndex        =   71
                  Top             =   750
                  Width           =   2265
                  _Version        =   196608
                  _ExtentX        =   3995
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
                  MaxValue        =   "70"
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
               Begin Threed.SSPanel pnl_CuoFin_MPr 
                  Height          =   315
                  Left            =   5820
                  TabIndex        =   72
                  Top             =   2730
                  Width           =   1065
                  _Version        =   65536
                  _ExtentX        =   1879
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_TipCam 
                  Height          =   315
                  Left            =   8910
                  TabIndex        =   73
                  Top             =   2400
                  Width           =   2265
                  _Version        =   65536
                  _ExtentX        =   3995
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_IntGra 
                  Height          =   315
                  Left            =   2610
                  TabIndex        =   74
                  Top             =   2730
                  Width           =   1365
                  _Version        =   65536
                  _ExtentX        =   2408
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_ValTot 
                  Height          =   315
                  Left            =   2610
                  TabIndex        =   75
                  Top             =   90
                  Width           =   1365
                  _Version        =   65536
                  _ExtentX        =   2408
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
                  Alignment       =   4
               End
               Begin EditLib.fpDoubleSingle ipp_ValEst 
                  Height          =   315
                  Left            =   3930
                  TabIndex        =   76
                  Top             =   420
                  Width           =   1125
                  _Version        =   196608
                  _ExtentX        =   1984
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
                  ButtonStyle     =   0
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
                  Text            =   "0.00"
                  DecimalPlaces   =   2
                  DecimalPoint    =   "."
                  FixedPoint      =   -1  'True
                  LeadZero        =   0
                  MaxValue        =   "9000000"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
                  IncInt          =   1
                  IncDec          =   1
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
               Begin Threed.SSPanel pnl_CuoIni 
                  Height          =   315
                  Left            =   2610
                  TabIndex        =   77
                  Top             =   750
                  Width           =   1365
                  _Version        =   65536
                  _ExtentX        =   2408
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_FmvBbp 
                  Height          =   315
                  Left            =   3930
                  TabIndex        =   78
                  Top             =   1080
                  Width           =   1125
                  _Version        =   65536
                  _ExtentX        =   1984
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_MtoBMS 
                  Height          =   315
                  Left            =   3930
                  TabIndex        =   79
                  Top             =   1410
                  Width           =   1125
                  _Version        =   65536
                  _ExtentX        =   1984
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_ValGas 
                  Height          =   315
                  Left            =   2610
                  TabIndex        =   80
                  Top             =   2070
                  Width           =   1365
                  _Version        =   65536
                  _ExtentX        =   2408
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_TotPre 
                  Height          =   315
                  Left            =   2610
                  TabIndex        =   81
                  Top             =   2400
                  Width           =   1365
                  _Version        =   65536
                  _ExtentX        =   2408
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
                  Alignment       =   4
               End
               Begin Threed.SSCheck chk_Gastos 
                  Height          =   285
                  Left            =   150
                  TabIndex        =   82
                  Top             =   2100
                  Width           =   1785
                  _Version        =   65536
                  _ExtentX        =   3149
                  _ExtentY        =   503
                  _StockProps     =   78
                  Caption         =   "Incluye Gastos"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSPanel pnl_MtoPre 
                  Height          =   315
                  Left            =   2610
                  TabIndex        =   83
                  Top             =   1740
                  Width           =   1365
                  _Version        =   65536
                  _ExtentX        =   2408
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
                  Alignment       =   4
               End
               Begin EditLib.fpDoubleSingle ipp_ComVta 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   116
                  Top             =   420
                  Width           =   1065
                  _Version        =   196608
                  _ExtentX        =   1879
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
                  ButtonStyle     =   0
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
                  Text            =   "0.00"
                  DecimalPlaces   =   2
                  DecimalPoint    =   "."
                  FixedPoint      =   -1  'True
                  LeadZero        =   0
                  MaxValue        =   "9000000000"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
                  IncInt          =   1
                  IncDec          =   1
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
               Begin EditLib.fpDoubleSingle ipp_ApoPro 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   117
                  Top             =   1080
                  Width           =   1065
                  _Version        =   196608
                  _ExtentX        =   1879
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
                  ButtonStyle     =   0
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
                  Text            =   "0.00"
                  DecimalPlaces   =   2
                  DecimalPoint    =   "."
                  FixedPoint      =   -1  'True
                  LeadZero        =   0
                  MaxValue        =   "9000000000"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
                  IncInt          =   1
                  IncDec          =   1
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
               Begin Threed.SSPanel pnl_CuoFin_Sol 
                  Height          =   315
                  Left            =   5820
                  TabIndex        =   118
                  Top             =   2400
                  Width           =   1065
                  _Version        =   65536
                  _ExtentX        =   1879
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
                  Alignment       =   4
               End
               Begin EditLib.fpDoubleSingle ipp_MtoAFP 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   119
                  Top             =   1410
                  Width           =   1065
                  _Version        =   196608
                  _ExtentX        =   1879
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
                  ButtonStyle     =   0
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
                  Text            =   "0.00"
                  DecimalPlaces   =   2
                  DecimalPoint    =   "."
                  FixedPoint      =   -1  'True
                  LeadZero        =   0
                  MaxValue        =   "9000000"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
                  IncInt          =   1
                  IncDec          =   1
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
               Begin VB.Label lbl_TipMon 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   225
                  Index           =   5
                  Left            =   5460
                  TabIndex        =   84
                  Top             =   2790
                  Width           =   285
               End
               Begin VB.Label lbl_TipMon 
                  Alignment       =   1  'Right Justify
                  Caption         =   "US$"
                  Height          =   285
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   115
                  Top             =   810
                  Width           =   375
               End
               Begin VB.Label Label36 
                  AutoSize        =   -1  'True
                  Caption         =   "Interés Capitalizado:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   150
                  TabIndex        =   114
                  Top             =   2790
                  Width           =   1755
               End
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo de Cambio"
                  Height          =   315
                  Left            =   7050
                  TabIndex        =   113
                  Top             =   2460
                  Width           =   1110
               End
               Begin VB.Label Label46 
                  Caption         =   "Préstamo (inc. gastos)"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   150
                  TabIndex        =   112
                  Top             =   2460
                  Width           =   2085
               End
               Begin VB.Label Label28 
                  Caption         =   "Mto. AFP (25%)"
                  Height          =   285
                  Left            =   300
                  TabIndex        =   111
                  Top             =   1470
                  Width           =   1275
               End
               Begin VB.Label Label23 
                  Caption         =   "Mto. BMS:"
                  Height          =   285
                  Left            =   2730
                  TabIndex        =   110
                  Top             =   1470
                  Width           =   915
               End
               Begin VB.Label Label24 
                  Caption         =   "Bonos:"
                  Height          =   285
                  Left            =   2730
                  TabIndex        =   109
                  Top             =   1140
                  Width           =   915
               End
               Begin VB.Label Label44 
                  Caption         =   "Cuota Inicial"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   150
                  TabIndex        =   108
                  Top             =   810
                  Width           =   1785
               End
               Begin VB.Label lbl_TipMon 
                  Alignment       =   1  'Right Justify
                  Caption         =   "US$"
                  Height          =   285
                  Index           =   4
                  Left            =   2160
                  TabIndex        =   107
                  Top             =   2460
                  Width           =   375
               End
               Begin VB.Label Label43 
                  Caption         =   "Valor Estac."
                  Height          =   285
                  Left            =   2730
                  TabIndex        =   106
                  Top             =   480
                  Width           =   915
               End
               Begin VB.Label Label42 
                  Caption         =   "Valor Total Vivienda"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   150
                  TabIndex        =   105
                  Top             =   150
                  Width           =   1785
               End
               Begin VB.Label Label33 
                  AutoSize        =   -1  'True
                  Caption         =   "Tasa Especial:"
                  Height          =   315
                  Left            =   7050
                  TabIndex        =   104
                  Top             =   2130
                  Width           =   1575
               End
               Begin VB.Label Label32 
                  Caption         =   "Cuotas Dobles:"
                  Height          =   315
                  Left            =   7050
                  TabIndex        =   103
                  Top             =   150
                  Width           =   1575
               End
               Begin VB.Label lbl_PorIni 
                  Alignment       =   1  'Right Justify
                  Caption         =   "(00.00%)"
                  Height          =   285
                  Left            =   4020
                  TabIndex        =   102
                  Top             =   810
                  Width           =   705
               End
               Begin VB.Label Label6 
                  Caption         =   "Interés Capitalizado PG:"
                  Height          =   315
                  Left            =   12120
                  TabIndex        =   101
                  Top             =   3090
                  Width           =   1785
               End
               Begin VB.Label Label9 
                  Caption         =   "Compañía Seguros:"
                  Height          =   315
                  Left            =   7050
                  TabIndex        =   100
                  Top             =   1470
                  Width           =   1575
               End
               Begin VB.Label Label10 
                  Caption         =   "Tipo Cambio "
                  Height          =   315
                  Left            =   12120
                  TabIndex        =   99
                  Top             =   2760
                  Width           =   1785
               End
               Begin VB.Label Label13 
                  Caption         =   "Cuota M. Prést.:"
                  Height          =   225
                  Left            =   4320
                  TabIndex        =   98
                  Top             =   2790
                  Width           =   1185
               End
               Begin VB.Label Label16 
                  Caption         =   "Cuota Final:"
                  Height          =   315
                  Left            =   4320
                  TabIndex        =   97
                  Top             =   2460
                  Width           =   855
               End
               Begin VB.Label Label27 
                  Caption         =   "Monto Préstamo:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   150
                  TabIndex        =   96
                  Top             =   1800
                  Width           =   2085
               End
               Begin VB.Label Label35 
                  Caption         =   "Valor Inmueble:"
                  Height          =   285
                  Left            =   300
                  TabIndex        =   95
                  Top             =   480
                  Width           =   1275
               End
               Begin VB.Label Label26 
                  Caption         =   "Aporte Propio:"
                  Height          =   285
                  Left            =   300
                  TabIndex        =   94
                  Top             =   1140
                  Width           =   1275
               End
               Begin VB.Label Label29 
                  Caption         =   "Plazo: (años)"
                  Height          =   315
                  Left            =   7050
                  TabIndex        =   93
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.Label Label25 
                  Caption         =   "Per. Gracia: (meses)"
                  Height          =   315
                  Left            =   7050
                  TabIndex        =   92
                  Top             =   810
                  Width           =   1575
               End
               Begin VB.Label Label22 
                  Caption         =   "Tipo Seg. Desgrav.:"
                  Height          =   315
                  Left            =   7050
                  TabIndex        =   91
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.Label Label18 
                  Caption         =   "Día de Pago:"
                  Height          =   315
                  Left            =   7050
                  TabIndex        =   90
                  Top             =   1140
                  Width           =   1575
               End
               Begin VB.Label lbl_TipMon 
                  Alignment       =   1  'Right Justify
                  Caption         =   "US$"
                  Height          =   285
                  Index           =   0
                  Left            =   2160
                  TabIndex        =   89
                  Top             =   150
                  Width           =   375
               End
               Begin VB.Label lbl_TipMon 
                  Alignment       =   1  'Right Justify
                  Caption         =   "US$"
                  Height          =   285
                  Index           =   2
                  Left            =   2160
                  TabIndex        =   88
                  Top             =   1800
                  Width           =   375
               End
               Begin VB.Label lbl_TipMon 
                  Alignment       =   1  'Right Justify
                  Caption         =   "US$"
                  Height          =   285
                  Index           =   3
                  Left            =   2160
                  TabIndex        =   87
                  Top             =   2130
                  Width           =   375
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   315
                  Left            =   5460
                  TabIndex        =   86
                  Top             =   2460
                  Width           =   285
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   315
                  Left            =   8190
                  TabIndex        =   85
                  Top             =   2460
                  Width           =   375
               End
            End
            Begin Threed.SSPanel SSPanel12 
               Height          =   3255
               Left            =   -74940
               TabIndex        =   120
               Top             =   360
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   5741
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
         End
      End
   End
   Begin Threed.SSPanel pnl_ComVta_Sol 
      Height          =   315
      Left            =   12840
      TabIndex        =   121
      Top             =   7440
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "0.00  "
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
      Alignment       =   4
   End
   Begin Threed.SSPanel pnl_ComVta_Dol 
      Height          =   315
      Left            =   13650
      TabIndex        =   123
      Top             =   7440
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "0  "
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
      Alignment       =   4
   End
   Begin Threed.SSPanel pnl_ApoPro_Sol 
      Height          =   315
      Left            =   12840
      TabIndex        =   124
      Top             =   7800
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "0.00  "
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
      Alignment       =   4
   End
   Begin Threed.SSPanel pnl_ApoPro_Dol 
      Height          =   315
      Left            =   13650
      TabIndex        =   125
      Top             =   7800
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "0  "
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
      Alignment       =   4
   End
   Begin Threed.SSPanel pnl_MtoPre_Sol 
      Height          =   315
      Left            =   12840
      TabIndex        =   127
      Top             =   8160
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "0.00  "
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
      Alignment       =   4
   End
   Begin Threed.SSPanel pnl_MtoPre_Dol 
      Height          =   315
      Left            =   13650
      TabIndex        =   128
      Top             =   8160
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "0  "
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
      Alignment       =   4
   End
   Begin VB.Label Label19 
      Caption         =   "Valores S/.:"
      Height          =   285
      Left            =   11910
      TabIndex        =   129
      Top             =   8190
      Width           =   885
   End
   Begin VB.Label Label17 
      Caption         =   "Valores S/.:"
      Height          =   285
      Left            =   11910
      TabIndex        =   126
      Top             =   7830
      Width           =   885
   End
   Begin VB.Label Label21 
      Caption         =   "Valores S/.:"
      Height          =   285
      Left            =   11910
      TabIndex        =   122
      Top             =   7470
      Width           =   885
   End
End
Attribute VB_Name = "frm_Seg_SolHip_65"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_Arr_TNC_Cli()     As String
Dim l_Arr_TC_Cli()      As String
Dim l_Arr_TNC_Cof()     As String
Dim l_Arr_TC_Cof()      As String

Dim l_arr_CuoExt()      As moddat_tpo_Genera
Dim l_arr_DiaPag()      As moddat_tpo_Genera
Dim l_arr_ParPrd()      As moddat_tpo_Genera
Dim l_arr_EmpSeg()      As moddat_tpo_Genera
Dim l_arr_CliNCo()      As modcal_g_est_CuoCli
Dim l_dbl_TipCam        As Double
Dim l_dbl_TasInt        As Double
Dim l_dbl_CuoApr        As Double
Dim l_dbl_IntGra        As Double
Dim l_str_EmpSeg        As String
Dim l_int_CodMod        As Integer
Dim l_int_PryMcs        As Integer
Dim l_int_TipEva        As Integer
Dim l_int_MesAho        As Integer
Dim l_int_AprCon        As Integer
Dim l_dbl_MtoPre_Cal    As Double
Dim l_int_PlaAno_Cal    As Integer
Dim l_int_TipSeg_Cal    As Integer
Dim l_int_PerGra_Cal    As Integer
Dim l_int_CuoDbl_Cal    As Integer
Dim l_str_Moneda_Pre    As String
Dim l_str_CodCiu        As String
Dim l_dbl_BmsTas        As Double
Dim l_dbl_MPSMS         As Double
Dim l_str_CodPry        As String
Dim l_dbl_GasTas        As Double
Dim l_dbl_GasNot        As Double
Dim l_dbl_BloReg        As Double
Dim l_dbl_RegMin        As Double
Dim l_dbl_RegHip        As Double
Dim l_dbl_ImpITF        As Double

Private Sub chk_Gastos_Click(Value As Integer)
   If Not fs_Valida_FinGCi(moddat_g_str_NumSol, CDbl(pnl_ValGas.Caption)) Then
      chk_Gastos.Value = False
   End If
   
   If chk_Gastos.Value = False Then
      pnl_ValGas.Caption = "0.00 "
   End If
      
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
   Call fs_Calcul
End Sub

Private Sub chk_Gastos_LostFocus()
   Call fs_Calcular_GCierre
End Sub

'**************************************************************************************************
'* BOTONES
'**************************************************************************************************
Private Sub cmd_DatInm_Click()
   If moddat_g_int_InmIde = 1 Then
      moddat_g_int_FlgGrb = 2
   Else
      moddat_g_int_FlgGrb = 1
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Seg_SolHip_54.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Call gs_LimpiaGrid(grd_Listad(2))

      Screen.MousePointer = 11
      Call modmip_gs_DatInm(grd_Listad(3), False)
      Screen.MousePointer = 0
   
      moddat_g_int_InmIde = 1
   End If
End Sub

Private Sub cmd_CalCuo_Click()
Dim r_dbl_ValViv        As Double
Dim r_dbl_MtoPre        As Double
Dim r_int_TipVal_Des    As Integer
Dim r_dbl_Import_Des    As Double
Dim r_int_TipVal_Viv    As Integer
Dim r_dbl_Import_Viv    As Double
Dim r_dbl_Portes        As Double
Dim r_dbl_MtoCon        As Double
Dim r_dbl_MtoNCo        As Double
Dim r_dbl_PorCon        As Double
Dim r_dbl_TopCon        As Double
Dim r_dbl_NuePre        As Double
Dim r_dbl_Contad        As Integer
Dim r_dbl_ComCof        As Double
Dim r_dbl_TasCof        As Double

'variables nueva para la generacion del cronograma
Dim obj_Cronog          As Object
Dim int_Produc          As Integer
Dim int_CuoDbl          As Integer
Dim dbl_ValInm          As Double
Dim dbl_CuoIni          As Double
Dim dbl_MtoCon          As Double
Dim int_PlaPre          As Integer
Dim dbl_TasInt          As Double
Dim dbl_TasCof          As Double
Dim dbl_ComCof          As Double
Dim dat_FecDes          As Date
Dim int_DiaVct          As Integer
Dim int_PerGra          As Integer
Dim str_PriVct          As String
Dim dbl_Portes          As Double
Dim dbl_SegViv          As Double
Dim int_TipSDe          As Integer
Dim dbl_SegDes          As Double
Dim dbl_CuoMen          As Double
Dim dbl_CuoPbp          As Double
Dim dbl_IngReq          As Double

   If ipp_ComVta.Value = 0 Then
      MsgBox "Debe ingresar el Valor de Compra Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta)
      Exit Sub
   End If
   'If ipp_ApoPro.Value = 0 Then
   '   MsgBox "Debe ingresar el Aporte Propio.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(ipp_ApoPro)
   '   Exit Sub
   'End If
   If ipp_PlaAno.Value = 0 Then
      MsgBox "Debe ingresar el Plazo de Créditos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   If cmb_CuoDbl.ListIndex = -1 Then
      MsgBox "Debe seleccionar si desea cuota extraordinaria.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CuoDbl)
      Exit Sub
   End If
   If cmb_EmpSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa de Seguro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta)
      Exit Sub
   End If
   If cmb_SegDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro de Desgravamen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegDes)
      Exit Sub
   End If
   If cmb_TasEsp.ListIndex = -1 Then
      MsgBox "Debe seleccionar una tasa especial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TasEsp)
      Exit Sub
   End If
   If cmb_DiaPag.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Día de Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegDes)
      Exit Sub
   End If
   
   cmd_CalCuo.Enabled = False
   Screen.MousePointer = 11
   
   'inicializa variables
   r_dbl_Contad = 0
   r_dbl_ValViv = CDbl(ipp_ComVta.Text)
   r_dbl_MtoPre = CDbl(pnl_TotPre.Caption)
   r_dbl_Portes = 0
   r_dbl_ComCof = 0
   r_dbl_TasCof = 0
   l_dbl_IntGra = 0
   
   'Determina tasa y comision de cofide
   If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then
      r_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, CInt(ipp_PlaAno.Text))
      r_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, CInt(ipp_PlaAno.Text))
   ElseIf InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
      r_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, CInt(ipp_PlaAno.Text))
      r_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, CInt(ipp_PlaAno.Text))
   End If
   
   'Obtiene Tasa de Seguro de Desgravamen y Vivienda
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo, Format(cmb_SegDes.ItemData(cmb_SegDes.ListIndex), "000"), moddat_g_int_TipMon, r_dbl_MtoPre, r_int_TipVal_Des, r_dbl_Import_Des, cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex))
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo, 0, moddat_g_int_TipMon, r_dbl_ValViv, r_int_TipVal_Viv, r_dbl_Import_Viv, cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex))
   
   'Obtiene portes
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   Select Case moddat_g_str_CodPrd > 0
      Case InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 2
         int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
         dbl_ValInm = CDbl(pnl_ValTot.Caption)
         dbl_CuoIni = CDbl(pnl_CuoIni.Caption) - CDbl(pnl_ValGas.Caption)
         dbl_MtoCon = 0
         int_PlaPre = CInt(ipp_PlaAno.Text) * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = 0
         dbl_ComCof = 0
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = cmb_DiaPag.ItemData(cmb_DiaPag.ListIndex)
         int_PerGra = CInt(ipp_PerGra.Text)
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = CInt(cmb_SegDes.ItemData(cmb_SegDes.ListIndex)) - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calcula cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra datos
         pnl_CuoFin_MPr.Caption = Format(dbl_CuoMen, "###,###,##0.00") & " "
         If moddat_g_int_TipMon = 1 Then
            pnl_CuoFin_Sol.Caption = Format(CDbl(pnl_CuoFin_MPr.Caption), "###,##0.00") & " "
         Else
            pnl_CuoFin_Sol.Caption = Format(CDbl(pnl_CuoFin_MPr.Caption) * l_dbl_TipCam, "###,##0.00") & " "
         End If
         
      Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd)  '"003"
         'determina tipo
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
         dbl_ValInm = CDbl(pnl_ValTot.Caption)
         dbl_CuoIni = CDbl(pnl_CuoIni.Caption) - CDbl(pnl_ValGas.Caption)
         dbl_MtoCon = CDbl(pnl_TotPre.Caption) * (r_dbl_PorCon / 100)  'ipp_MtoPre.Value
         If dbl_MtoCon > r_dbl_TopCon Then dbl_MtoCon = r_dbl_TopCon
         int_PlaPre = CInt(ipp_PlaAno.Text) * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = 0
         dbl_ComCof = 0
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = cmb_DiaPag.ItemData(cmb_DiaPag.ListIndex)
         int_PerGra = CInt(ipp_PerGra.Text)
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = CInt(cmb_SegDes.ItemData(cmb_SegDes.ListIndex)) - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calcula cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra datos
         pnl_CuoFin_MPr.Caption = Format(dbl_CuoPbp, "###,###,##0.00") & " "
         pnl_CuoFin_Sol.Caption = Format(CDbl(pnl_CuoFin_MPr.Caption), "###,##0.00") & " "
         
      Case InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) Or InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd)
         r_dbl_TopCon = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         If CDbl(pnl_ComVta_Sol.Caption) > (50 * moddat_gf_Consulta_ParVal("001", "002")) Then
            r_dbl_TopCon = 5000
         End If
         
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
         dbl_ValInm = CDbl(pnl_ValTot.Caption)
         dbl_CuoIni = CDbl(pnl_CuoIni.Caption) - CDbl(pnl_ValGas.Caption)
         dbl_MtoCon = r_dbl_TopCon
         int_PlaPre = CInt(ipp_PlaAno.Text) * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = cmb_DiaPag.ItemData(cmb_DiaPag.ListIndex)
         int_PerGra = CInt(ipp_PerGra.Text)
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = CInt(cmb_SegDes.ItemData(cmb_SegDes.ListIndex)) - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calcula cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra datos
         pnl_CuoFin_MPr.Caption = Format(dbl_CuoPbp, "###,###,##0.00") & " "
         pnl_CuoFin_Sol.Caption = Format(CDbl(pnl_CuoFin_MPr.Caption), "###,##0.00") & " "
         
      Case InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 3
         int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
         dbl_ValInm = CDbl(pnl_ValTot.Caption)
         dbl_CuoIni = CDbl(pnl_CuoIni.Caption) - CDbl(pnl_ValGas.Caption)
         dbl_MtoCon = 0
         int_PlaPre = CInt(ipp_PlaAno.Text) * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = cmb_DiaPag.ItemData(cmb_DiaPag.ListIndex)
         int_PerGra = CInt(ipp_PerGra.Text)
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = CInt(cmb_SegDes.ItemData(cmb_SegDes.ListIndex)) - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra datos
         pnl_CuoFin_MPr.Caption = Format(dbl_CuoMen, "###,###,##0.00") & " "
         pnl_CuoFin_Sol.Caption = Format(CDbl(pnl_CuoFin_MPr.Caption), "###,##0.00") & " "
   End Select
   
   'calcula interes capitalizado
   If int_PerGra > 0 Then
      l_dbl_IntGra = Format(l_Arr_TNC_Cli(int_PerGra, 10) - CDbl(pnl_TotPre.Caption), "###,##0.00")  'ipp_MtoPre.Text
   End If
   
   pnl_IntGra.Caption = Format(l_dbl_IntGra, "###,##0.00") & " "
   Screen.MousePointer = 0
   cmd_CalCuo.Enabled = True
End Sub

Private Sub cmd_Aprueb_Click()
Dim r_dbl_ValMin_ComVta    As Double
Dim r_dbl_ValMax_ComVta    As Double
Dim r_dbl_PorMin_ApoPro    As Double
Dim r_dbl_PorMax_MtoPre    As Double
Dim r_dbl_ValMin_MtoPre    As Double
Dim r_dbl_ValMax_MtoPre    As Double
Dim r_int_DiaTra           As Integer
Dim r_str_Cadena           As String
Dim r_dbl_TCaDol           As Double
Dim r_int_EdaMax           As Integer
Dim r_int_EdaAct           As Integer
Dim r_dbl_ApoMin           As Double
Dim r_dbl_Ini_ApoMin       As Double
Dim r_dbl_PrcMin           As Double
Dim r_dbl_PrcMax           As Double
Dim r_dbl_Aho_ApoTp1       As Double
Dim r_int_FlgExc           As Integer
Dim r_str_Parame           As String
Dim r_rst_Genera           As ADODB.Recordset
Dim r_int_flgest           As Integer
Dim r_int_Resul            As Integer
Dim r_str_CodMod           As String
Dim r_str_CodPrd           As String
Dim r_str_DesMod           As String

   If moddat_g_int_InmIde = 2 Then
      MsgBox "El Cliente no ha registrado información del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_DatInm)
      Exit Sub
   End If
   
   r_int_flgest = 0
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT nvl((select a.solinm_flgest "
   r_str_Parame = r_str_Parame & "   FROM cre_solinm a "
   r_str_Parame = r_str_Parame & "  WHERE a.solinm_numsol = '" & moddat_g_str_NumSol & "'),0) AS solinm_flgest "
   r_str_Parame = r_str_Parame & "   FROM dual "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      r_int_flgest = r_rst_Genera!SOLINM_FLGEST
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
   
   If r_int_flgest = 1 Then
      If CDbl(ipp_ValEst.Text) = 0 Then
         MsgBox "Tiene que ingresar el valor del estacionamiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ValEst)
         Exit Sub
      End If
   Else
      If CDbl(ipp_ValEst.Text) <> 0 Then
         MsgBox "El valor del estacionamiento tiene que ser cero.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ValEst)
         Exit Sub
      End If
   End If
   
   'Edad Máxima del Cliente
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "012") Then
      r_int_EdaMax = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Buscar Parámetros en Productos
   r_dbl_ValMin_ComVta = 0
   r_dbl_ValMax_ComVta = 0
   r_dbl_PorMin_ApoPro = 0
   r_dbl_PorMax_MtoPre = 0
   r_dbl_ValMin_MtoPre = 0
   r_dbl_ValMax_MtoPre = 0
   r_dbl_PrcMin = 0
   r_dbl_PrcMax = 0
   
   'Para obtener Valor Máximo del Inmueble
   Select Case moddat_g_str_CodPrd > 0
      Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd)   '"001"  'En UIT (Mínimo y Máximo)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      Case InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd)  '"002", "011",   'En Montos
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "021") Then
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_Cantid
         End If
      
      Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd)   '"003"  'En UIT
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      Case InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd)  '"004"  'En UIT (Mínimo y Máximo)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      Case InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMin_ComVta = Format(r_dbl_ValMin_ComVta, "###,###,##0.00")
            r_dbl_ValMax_ComVta = Format(r_dbl_ValMax_ComVta, "###,###,##0.00")
         End If
   End Select

   'Para obtener % Mínimo de Aporte Propio
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "022") Then
      r_dbl_PorMin_ApoPro = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Para obtener % Máximo de Monto de Préstamo
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "023") Then
      r_dbl_PorMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Para obtener Monto Máximo de Préstamo
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then      '"002" "011"
      'En Monto Máximo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "024") Then
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'En Monto Mínimo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "026") Then
         r_dbl_ValMin_MtoPre = l_arr_ParPrd(1).Genera_Cantid
      End If
      
   ElseIf InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then '"001" "003" Then
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "023") Then
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid * moddat_gf_Consulta_ParVal("001", "002")
      End If
      
   ElseIf InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then  '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023" Then
      'Porcentaje para Valor Minimo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "024") Then
         r_dbl_PrcMin = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'Porcentaje para Valor Máximo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "025") Then
         r_dbl_PrcMax = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "023") Then
         r_dbl_ValMin_MtoPre = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMin / 100
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMax / 100
      End If
   End If
   
   'Validando Valor de Compra Venta
   If CDbl(pnl_ValTot.Caption) = 0 Then
      MsgBox "Debe ingresar el Valor de Compra-Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta)
      Exit Sub
   End If
   
   If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
      If CDbl(pnl_ComVta_Sol.Caption) < r_dbl_ValMin_ComVta Then
         MsgBox "El Valor de Compra-Venta no cubre el mínimo requerido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
      If CDbl(pnl_ComVta_Sol.Caption) > r_dbl_ValMax_ComVta Then
         MsgBox "El Valor de Compra-Venta excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
   ElseIf InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      If CDbl(pnl_ComVta_Sol.Caption) > r_dbl_ValMax_ComVta Then
         MsgBox "El Valor de Compra-Venta excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
   End If
   
   'Validando cuota inicial
   'If CDbl(ipp_ApoPro.Text) = 0 Then
   '   MsgBox "Debe ingresar el Aporte Propio.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(ipp_ApoPro)
   '   Exit Sub
   'End If
   
   If CDbl(pnl_CuoIni.Caption) > CDbl(pnl_ValTot.Caption) Then
      MsgBox "El Aporte Propio no puede ser mayor al Valor de Compra Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   
   If moddat_g_str_CodPrd = "023" Then
      If CDbl(Format(CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption) * 100, "###0.00")) < r_dbl_PorMin_ApoPro Then
         MsgBox "El Aporte Propio no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   Else
      If CDbl(Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP.Text)) / CDbl(pnl_ValTot.Caption) * 100, "###0.00")) < r_dbl_PorMin_ApoPro Then
         MsgBox "El Aporte Propio no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If
   
   If CDbl(pnl_MtoPre.Caption) / CDbl(pnl_ValTot.Caption) * 100 > r_dbl_PorMax_MtoPre Then
      MsgBox "El Aporte Propio no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If

   'Validando Monto de Préstamo
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
      If moddat_g_int_TipMon = 1 Then
         If CDbl(pnl_MtoPre_Sol.Caption) < r_dbl_ValMin_MtoPre Then
            MsgBox "El Monto del Préstamo no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
         If CDbl(pnl_MtoPre_Sol.Caption) > r_dbl_ValMax_MtoPre Then
            MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      Else
         If CDbl(pnl_MtoPre_Dol.Caption) < r_dbl_ValMin_MtoPre Then
            MsgBox "El Monto del Préstamo no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
         
         If CDbl(pnl_MtoPre_Dol.Caption) > r_dbl_ValMax_MtoPre Then
            MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
      
   ElseIf InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then
      If CDbl(pnl_MtoPre_Sol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
      
   ElseIf InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
      If CDbl(pnl_MtoPre_Sol.Caption) < r_dbl_ValMin_MtoPre Then
         MsgBox "El Monto del Préstamo no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
      
      If CDbl(pnl_MtoPre_Sol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If

   If CDbl(ipp_PlaAno.Text) = 0 Then
      MsgBox "Debe ingresar el Plazo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   If Not (CInt(ipp_PlaAno.Text) >= ipp_PlaAno.MinValue And CInt(ipp_PlaAno.Text) <= ipp_PlaAno.MaxValue) Then
      MsgBox "El Plazo indicado no se ajusta a los Parámetros permitidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   r_int_EdaAct = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Tit), date), 2))
   If r_int_EdaAct + CInt(ipp_PlaAno.Text) > r_int_EdaMax Then
      MsgBox "La Edad del Cliente más el Plazo del Préstamo excede el parámetro permitido. El Plazo máximo podría ser de " & CStr(r_int_EdaMax - r_int_EdaAct) & " años.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   If Not (CInt(ipp_PerGra.Value) >= CInt(ipp_PerGra.MinValue) And CInt(ipp_PerGra.Value) <= CInt(ipp_PerGra.MaxValue)) Then
      MsgBox "El Período de Gracia no se ajusta al parámetro permitido. El Plazo máximo podría ser de " & CStr(ipp_PerGra.MaxValue) & " meses.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerGra)
      Exit Sub
   End If
   If cmb_EmpSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa de Seguros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EmpSeg)
      Exit Sub
   End If
   If cmb_SegDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro de Desgravamen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegDes)
      Exit Sub
   End If
   If cmb_TasEsp.ListIndex = -1 Then
      MsgBox "Debe seleccionar una tasa especial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TasEsp)
      Exit Sub
   End If
   If moddat_g_int_EstCiv <> 2 And moddat_g_int_EstCiv <> 5 Then
      If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) = 12 Then
         MsgBox "El Cliente no requiere tomar Seguro de Desgravamen Mancomunado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SegDes)
         Exit Sub
      End If
   End If
   
   'Si cliente complementa Renta
   If moddat_g_int_ComRta = 1 Then
      If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) <> 12 Then
         MsgBox "El Cliente presenta Complemento de Renta debe seleccionar el Seguro de Desgravamen Mancomunado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SegDes)
         Exit Sub
      End If
   End If
   
   If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) = 12 Then
      r_int_EdaAct = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Cyg), date), 2))
      
      If r_int_EdaAct + CInt(ipp_PlaAno.Text) > r_int_EdaMax Then
         MsgBox "La Edad del Cónyuge más el Plazo del Préstamo excede el parámetro permitido. El Plazo máximo podría ser de " & CStr(r_int_EdaMax - r_int_EdaAct) & " años.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PlaAno)
         Exit Sub
      End If
   End If
   
   If cmb_DiaPag.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Día de Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DiaPag)
      Exit Sub
   End If
   
   If l_int_TipEva = 1 Then
      If moddat_g_str_CodPrd = "023" Then
         r_dbl_ApoMin = CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption) * 100
      Else
         r_dbl_ApoMin = (CDbl(ipp_ApoPro.Text) + CDbl(Me.ipp_MtoAFP.Text)) / CDbl(pnl_ValTot.Caption) * 100
      End If
      
      'Validando que Clientes de Provincias cumplan con Aporte Inicial mínimo
      If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "025") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "Cliente de Provincias. El Aporte Inicial es menor al Aporte Inicial mínimo requerido. (" & CStr(r_dbl_Ini_ApoMin) & "%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
   End If
   
   'Ahorro Programado
   If l_int_TipEva = 2 Then
      If moddat_g_str_CodPrd = "023" Then
         r_dbl_ApoMin = CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption) * 100
      Else
         r_dbl_ApoMin = (CDbl(ipp_ApoPro.Text) + CDbl(Me.ipp_MtoAFP.Text)) / CDbl(pnl_ValTot.Caption) * 100
      End If
      
      r_dbl_Aho_ApoTp1 = 0
      
      'Clientes de Provincia no tienen acceso a este Tipo de Evaluación
      'If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
      '   MsgBox "Este tipo de Evaluación sólo está permitida para clientes que residen en Lima Metropolitana o Callao.", vbExclamation, modgen_g_str_NomPlt
      '   Exit Sub
      'End If
      
      If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then
         If r_dbl_ApoMin < 20 Then
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (20%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
            r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
         End If
         If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
            MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
      ElseIf InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then  '"003"
         If r_dbl_ApoMin >= 20 And r_dbl_ApoMin < 30 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "013") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
         ElseIf r_dbl_ApoMin >= 30 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "014") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         Else
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (20%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
         
      ElseIf InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Then '"004"
         If r_dbl_ApoMin >= 10 And r_dbl_ApoMin < 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "011") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         ElseIf r_dbl_ApoMin >= 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         Else
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (10%).", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
      ElseIf InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then
         If r_dbl_ApoMin >= 10 And r_dbl_ApoMin < 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "011") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         ElseIf r_dbl_ApoMin >= 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If l_int_MesAho < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         Else
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (10%).", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
      End If
   End If
   
   'Cuota Inicial 30%-35%
   If l_int_TipEva = 3 Then
      If moddat_g_str_CodPrd = "023" Then
         r_dbl_ApoMin = CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption) * 100
      Else
         r_dbl_ApoMin = (CDbl(ipp_ApoPro.Text) + CDbl(Me.ipp_MtoAFP.Text)) / CDbl(pnl_ValTot.Caption) * 100
      End If
      
      If l_int_PryMcs = 1 Then
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "001") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "El Aporte Inicial es menor al Aporte Inicial mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      Else
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "002") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "El Aporte Inicial es menor al Aporte Inicial mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
   End If
   
   'Cuota Inicial 50% Sin Evaluación
   If l_int_TipEva = 4 Then
      If moddat_g_str_CodPrd = "023" Then
         r_dbl_ApoMin = CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption) * 100
      Else
         r_dbl_ApoMin = (CDbl(ipp_ApoPro.Text) + CDbl(Me.ipp_MtoAFP.Text)) / CDbl(pnl_ValTot.Caption) * 100
      End If
      
      r_dbl_Ini_ApoMin = 0
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "054", "001") Then
         r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
      End If
      If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
         MsgBox "El Aporte Inicial es menor al Aporte Inicial mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If
   
   If l_int_TipEva <> 1 Then
      If CDbl(pnl_TotPre.Caption) > l_dbl_MtoPre_Cal Then
         MsgBox "El Monto del Préstamo aprobado es de " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & Format(l_dbl_MtoPre_Cal, "###,###0.00") & ". Este Tipo de Evaluación no permite cambios de condiciones de Aprobación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
      If CInt(ipp_PlaAno.Text) <> l_int_PlaAno_Cal Then
         MsgBox "El Plazo del Préstamo aprobado es de " & CStr(l_int_PlaAno_Cal) & " años. Este Tipo de Evaluación no permite cambios de condiciones de Aprobación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PlaAno)
         Exit Sub
      End If
      If CInt(ipp_PerGra.Text) <> l_int_PerGra_Cal Then
         MsgBox "El Período de Gracia del Préstamo aprobado es de " & CStr(l_int_PerGra_Cal) & " meses. Este Tipo de Evaluación no permite cambios de condiciones de Aprobación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PerGra)
         Exit Sub
      End If
      If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) <> l_int_TipSeg_Cal Then
         MsgBox "El Tipo de Seguro no coincide con el Tipo de Seguro aprobado. Este Tipo de Evaluación no permite cambios de condiciones de Aprobación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SegDes)
         Exit Sub
      End If
   End If
   
   'Validando la Cuota
   If CDbl(pnl_CuoFin_MPr.Caption) = 0 Then
      MsgBox "Debe calcular la Cuota a Pagar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_CalCuo)
      Exit Sub
   End If
   
   r_int_FlgExc = 1
   
   If CDbl(pnl_CuoFin_Sol.Caption) > l_dbl_CuoApr Then
      If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Or (InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 And (l_int_TipEva = 3 Or l_int_TipEva = 4)) Then
         If modgen_g_int_TipUsu = 20200 Or modgen_g_int_TipUsu = 20220 Or modgen_g_int_TipUsu = 1000 Then
            If MsgBox(moddat_g_str_Msje01 & " " & "¿Desea aprobar la operación?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            End If
            r_int_FlgExc = 2
         Else
            'Solicitar código de autorización
            MsgBox moddat_g_str_Msje01 & " " & "No se puede aprobar la operación.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmd_CalCuo)
            Exit Sub
         End If
      Else
         MsgBox moddat_g_str_Msje01 & " " & "No se puede aprobar la operación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_CalCuo)
         Exit Sub
      End If
   End If
   
   '******************************************************************************************************************
   r_str_CodMod = ""
   r_str_CodPrd = ""
   r_str_DesMod = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "  SELECT SOLMAE_CODPRD, SOLMAE_CODMOD"
   r_str_Parame = r_str_Parame & "    FROM CRE_SOLMAE "
   r_str_Parame = r_str_Parame & "   WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      r_str_CodMod = Trim(r_rst_Genera!SOLMAE_CODMOD)
      r_str_CodPrd = Trim(r_rst_Genera!SOLMAE_CODPRD)
      r_str_DesMod = moddat_gf_Buscar_NomMod(Trim(r_str_CodPrd), r_str_CodMod)
   End If
   
   'VALIDACION DE PARAMETROS ASOCIADOS A LOS GASTOS DE CIERRE
   If InStr(r_str_DesMod, "TERMINADO") = 0 Then
      
      'Valida los Gastos de Cierre
      r_int_Resul = gf_Valida_GastoCierre(r_str_CodPrd, l_str_CodPry)
      
      If r_int_Resul = 1 Then
         MsgBox "El proyecto asociado no tiene empresa de peritaje asignado, favor actualizar información en la plataforma de Operaciones (area operativa).", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
         Exit Sub
      ElseIf r_int_Resul = 2 Then
         MsgBox "El proyecto asociado no tiene notaría asignada, favor actualizar información en la plataforma de Operaciones (area legal).", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
         Exit Sub
      ElseIf r_int_Resul = 3 Then
         If MsgBox("La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor coordinar con el área legal la actualización de la información en caso contrario no se generaran los gastos de cierre." & vbCrLf & "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            chk_Gastos.Value = False
            Exit Sub
         End If
      ElseIf r_int_Resul = 4 Then
         MsgBox "La empresa de peritaje asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar con el área de operaciones la actualizacion de la información.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
         Exit Sub
      End If
   
      'Valida que los Gastos de Cierre se puedan actualizar
      If Not fs_Valida_FinGCi(moddat_g_str_NumSol, CDbl(pnl_ValGas.Caption)) Then
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'Registrando Excepción
   If r_int_FlgExc = 2 Then
      Call fs_RegExc
   End If

   r_dbl_TCaDol = moddat_gf_Obtiene_TipCam(1, 2)
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 31)))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 31, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 31, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando Datos de Aprobación en Solicitud de Crédito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis

      g_str_Parame = "USP_CRE_SOLMAE_ACECLI ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_SegDes.ItemData(cmb_SegDes.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_DiaPag.ItemData(cmb_DiaPag.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PlaAno.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PlaAno.Value * 12) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PerGra.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ComVta_Dol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ComVta_Sol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ApoPro_Dol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoIni.Caption)) & ", " 'pnl_ApoPro_Sol
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre_Dol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre_Sol.Caption)) & ", "
      If moddat_g_int_TipMon = 1 Then
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre_Sol.Caption)) & ", "
      Else
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre_Dol.Caption)) & ", "
      End If
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoFin_Sol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoFin_MPr.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_TCaDol) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TipCam.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IntGra) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoBMS.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ComVta.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEst.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoAFP.Text)) & ", "
      'g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ValGas.Caption)) & ", "

      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Inserta Nueva Instancia de Evaluación
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 32) Then
      Exit Sub
   End If
      
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 32, 11, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando en Tabla de Créditos
   If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, 32) Then
      Exit Sub
   End If
   
   'Valida y Actualiza Financiamiento
   If Not fs_Actualiza_FinGCi(moddat_g_str_NumSol, CDbl(pnl_ValGas.Caption)) Then
      Exit Sub
   End If
   
   'Enviando Correo Electrónico
   modgen_g_str_Mail_Asunto = moddat_gf_Consulta_ParDes("002", "31") & " - APROBACION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   
   Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False)
   
   MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_con_AteCli
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Rechaz_Click()
Dim r_int_DiaTra     As Integer
Dim r_str_CodIns     As String
Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = 31
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 31)))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 31, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 31, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      'Actualizando Rechazo en Tabla de Créditos
      If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      modgen_g_str_Mail_Asunto = moddat_gf_Consulta_ParDes("002", "31") & " - RECHAZO (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_gf_Consulta_ParDes("003", moddat_g_int_MotRec)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False)
   
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_con_AteCli
      moddat_g_int_FlgAct = 2
      Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

'**************************************************************************************************
'* FORM
'**************************************************************************************************
Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Limpia
   
   'Obteniendo Tipo de Cambio de Moneda del Préstamo
   l_dbl_TipCam = 0
   l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, 2)
   pnl_TipCam.Caption = Format(l_dbl_TipCam, "###,##0.0000") & " "
   
   If l_dbl_TipCam = 0 Then
      MsgBox "El tipo de cambio no esta registrado.", vbInformation, modgen_g_con_AteCli
      cmd_DatInm.Enabled = False
      cmd_CalCuo.Enabled = False
      cmd_Aprueb.Enabled = False
      cmd_Rechaz.Enabled = False
      
      ipp_ComVta.Enabled = False
      ipp_ValEst.Enabled = False
      ipp_ApoPro.Enabled = False
      ipp_MtoAFP.Enabled = False
      chk_Gastos.Enabled = False
      cmb_CuoDbl.Enabled = False
      ipp_PlaAno.Enabled = False
      ipp_PerGra.Enabled = False
      cmb_DiaPag.Enabled = False
      cmb_EmpSeg.Enabled = False
      cmb_SegDes.Enabled = False
      cmb_TasEsp.Enabled = False
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Buscar Información de la Solicitud
   moddat_g_str_FecNac_Tit = ""
   moddat_g_str_FecNac_Cyg = ""
   moddat_g_int_RegCyg = 0
   moddat_g_int_EstCiv = 0
   moddat_g_int_ComRta = 0
   
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Información del Cónyuge
   Call modmip_gs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(2))         'Buscar Información del Apoderado
   Call modmip_gs_DatInm(grd_Listad(3), False)                                            'Buscar Información del Inmueble
   Call modmip_gs_DatCre(grd_Listad(4), r_arr_Mtz)                                        'Buscar Información del Crédito
   
   txt_ObsSol.Text = r_arr_Mtz(0).DatCom_Observ
   lbl_PorIni.Caption = "(0.00%)"
   If r_arr_Mtz(0).DatCom_ComVta_Mon > 0 Then
      If r_arr_Mtz(0).DatCom_ComVta_Mon = 1 Then
         pnl_ValTot.Caption = Format(CDbl(r_arr_Mtz(0).DatCom_ComVta_Sol), "##,###,##0.00") & "  "
         ipp_ComVta.Value = IIf(r_arr_Mtz(0).DatCom_MtoInm_Sol = 0, CDbl(r_arr_Mtz(0).DatCom_ComVta_Sol), CDbl(r_arr_Mtz(0).DatCom_MtoInm_Sol))
         ipp_ValEst.Value = r_arr_Mtz(0).DatCom_MtoEst_Sol
         pnl_CuoIni.Caption = Format(CDbl(r_arr_Mtz(0).DatCom_ApoPro_Sol), "##,###,##0.00") & "  "
         ipp_ApoPro.Value = r_arr_Mtz(0).DatCom_ApoPro_Sol - (r_arr_Mtz(0).DatCom_FmvBbp_Sol + r_arr_Mtz(0).DatCom_MefPbp_Sol + r_arr_Mtz(0).DatCom_MtoAFP_Sol + r_arr_Mtz(0).DatCom_MtoBMS_Sol)
         pnl_FmvBbp.Caption = Format(CDbl(r_arr_Mtz(0).DatCom_FmvBbp_Sol) + CDbl(r_arr_Mtz(0).DatCom_MefPbp_Sol), "##,###,##0.00") & "  "
         ipp_MtoAFP.Value = r_arr_Mtz(0).DatCom_MtoAFP_Sol
         pnl_MtoBMS.Caption = Format(CDbl(r_arr_Mtz(0).DatCom_MtoBMS_Sol), "##,###,##0.00") & "  "
         pnl_TotPre.Caption = r_arr_Mtz(0).DatCom_MtoPre_Sol
      Else
         ipp_ComVta.Value = r_arr_Mtz(0).DatCom_ComVta_Dol
         ipp_ApoPro.Value = r_arr_Mtz(0).DatCom_ComVta_Dol - r_arr_Mtz(0).DatCom_MtoPre_Cal
      End If
      pnl_MtoPre.Caption = Format(CDbl(r_arr_Mtz(0).DatCom_PreMto_Sol), "##,###,##0.00") & "  "
      pnl_ValGas.Caption = Format(CDbl(r_arr_Mtz(0).DatCom_MtoGCi_Sol), "##,###,##0.00") & "  "
      If CDbl(r_arr_Mtz(0).DatCom_MtoGCi_Sol) > 0 Then
         chk_Gastos.Value = True
      Else
         chk_Gastos.Value = False
      End If
      pnl_TotPre.Caption = Format(CDbl(r_arr_Mtz(0).DatCom_MtoPre_Cal), "##,###,##0.00") & "  "
      
      
      Call fs_Calcular_GCierre
      Call fs_Calcular_Prestamo
      Call fs_Calcul
      
      ipp_PlaAno.Value = r_arr_Mtz(0).DatCom_PlaAno_Cal
      ipp_PerGra.Value = r_arr_Mtz(0).DatCom_PerGra_Cal
      Call gs_BuscarCombo_Item(cmb_CuoDbl, r_arr_Mtz(0).DatCom_CuoExt_Cal)
      cmb_EmpSeg.ListIndex = gf_Busca_Arregl(l_arr_EmpSeg, r_arr_Mtz(0).DatCom_EsgDes) - 1
      Call moddat_gs_Carga_TipSeg(cmb_SegDes, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo)
      
      Call gs_BuscarCombo_Item(cmb_SegDes, r_arr_Mtz(0).DatCom_TipSeg_Cal)
      Call gs_BuscarCombo_Item(cmb_TasEsp, r_arr_Mtz(0).DatCom_TasEsp)
      Call gs_BuscarCombo_Item(cmb_DiaPag, r_arr_Mtz(0).DatCom_DiaPag)
      
      l_dbl_BmsTas = r_arr_Mtz(0).DatCom_BmsTas
      l_str_EmpSeg = r_arr_Mtz(0).DatCom_EsgDes
      l_int_MesAho = r_arr_Mtz(0).DatCom_MesAho
   End If
   l_int_TipEva = r_arr_Mtz(0).DatCom_TipEva
   l_dbl_TasInt = r_arr_Mtz(0).DatCom_TasInt
   moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
   moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
   
   Call modmip_gs_EvaCre(grd_Listad(5))
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)
   Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   Call fs_EvaCre
   
   'Periodo de Gracia
   If l_int_CodMod > 0 Then
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "008", Format(l_int_CodMod, "000")) Then
         ipp_PerGra.MinValue = moddat_g_arr_Genera(1).Genera_ValMin
         ipp_PerGra.MaxValue = moddat_g_arr_Genera(1).Genera_ValMax
      End If
   End If
   
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
   
   'Si tiene perfil de consejerop no puede modificar los datos
   If modgen_g_int_TipUsu = 20121 Then
      MsgBox "Perfil de consejero hipotecario no esta habilitado para modificar datos.", vbExclamation, modgen_g_str_NomPlt
      ipp_ComVta.Enabled = False
      ipp_ApoPro.Enabled = False
      pnl_TotPre.Enabled = False 'ipp_MtoPre
      ipp_PlaAno.Enabled = False
      ipp_PerGra.Enabled = False
      cmb_CuoDbl.Enabled = False
      cmb_EmpSeg.Enabled = False
      cmb_SegDes.Enabled = False
      cmb_DiaPag.Enabled = False
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

'**************************************************************************************************
'* PROCEDIMIENTOS Y FUNCIONES
'**************************************************************************************************
Private Sub fs_Inicia()
Dim r_int_Contad     As Integer
   
   'Plazo de Crédito
   If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub) Then
      ipp_PlaAno.MinValue = moddat_g_arr_Genera(1).Genera_PlzMin
      ipp_PlaAno.MaxValue = moddat_g_arr_Genera(1).Genera_PlzMax
   End If
   
   'Periodo de Gracia
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "008", "002") Then
      ipp_PerGra.MinValue = moddat_g_arr_Genera(1).Genera_ValMin
      ipp_PerGra.MaxValue = moddat_g_arr_Genera(1).Genera_ValMax
   End If
   
   'Carga combos
   Call moddat_gs_Carga_ParSubPrd_ComboItem(cmb_DiaPag, moddat_g_str_CodPrd, moddat_g_str_CodSub, "009")
   Call moddat_gs_Carga_EmpSeg(cmb_EmpSeg, l_arr_EmpSeg)
   Call moddat_gs_Carga_LisIte_Combo(cmb_CuoDbl, 1, "277")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TasEsp, 1, "522")
   
   'Busca descripcion de la moneda
   l_str_Moneda_Pre = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon))
   
   'Inicializando Grid de Cliente y de Cónyuge
   For r_int_Contad = 0 To 5
      grd_Listad(r_int_Contad).ColWidth(0) = 3000
      grd_Listad(r_int_Contad).ColWidth(1) = 7940
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad
   
   If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub) Then
      DoEvents
   End If

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
   
   'Busca Proyecto
   l_str_CodPry = gf_Obtener_Proyec(moddat_g_str_NumSol)
End Sub

Private Sub fs_Limpia()
   ipp_ComVta.Value = 0
   ipp_ApoPro.Value = 0
   pnl_TotPre.Caption = 0 'ipp_MtoPre.Value
   ipp_PlaAno.Value = ipp_PlaAno.MinValue
   ipp_PerGra.Value = 0
   cmb_CuoDbl.ListIndex = -1
   cmb_EmpSeg.ListIndex = -1
   cmb_SegDes.Clear
   cmb_DiaPag.ListIndex = -1
   
   pnl_CuoFin_Sol.Caption = "0.00 "
   pnl_CuoFin_MPr.Caption = "0.00 "
   pnl_TipCam.Caption = "0.0000 "
   pnl_IntGra.Caption = "0.00 "
   
   lbl_TipMon(0).Caption = l_str_Moneda_Pre
   lbl_TipMon(1).Caption = l_str_Moneda_Pre
   lbl_TipMon(2).Caption = l_str_Moneda_Pre
   lbl_TipMon(3).Caption = l_str_Moneda_Pre
   lbl_TipMon(4).Caption = l_str_Moneda_Pre
   lbl_TipMon(5).Caption = l_str_Moneda_Pre
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      If p_Indice = 0 Then
         moddat_g_str_FecNac_Tit = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      Else
         moddat_g_str_FecNac_Cyg = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      End If
      
      If p_Indice = 0 Then
         If g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
            moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
         End If
         moddat_g_int_EstCiv = g_rst_Princi!DATGEN_ESTCIV
         moddat_g_int_RegCyg = g_rst_Princi!DATGEN_REGCYG
         l_str_CodCiu = g_rst_Princi!DATGEN_CODCIU
      End If
      
      If p_Indice = 0 Then
         moddat_g_str_UbiGeo = Left(Format(g_rst_Princi!DatGen_Ubigeo, "000000"), 4)
      End If
      
      If p_Indice = 1 Then
         moddat_g_int_ComRta = 0
         If g_rst_Princi!DATGEN_ACTECO = 1 Then
            moddat_g_int_ComRta = 1
         End If
         
         If CInt(l_str_CodCiu) <> 7522 And CInt(l_str_CodCiu) <> 7523 Then
            l_str_CodCiu = g_rst_Princi!DATGEN_CODCIU
         End If
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Calcul()
   If moddat_g_int_TipMon = 1 Then
      pnl_ComVta_Sol.Caption = pnl_ValTot.Caption
      pnl_ApoPro_Sol.Caption = pnl_CuoIni.Caption
      pnl_MtoPre_Sol.Caption = pnl_TotPre.Caption
      pnl_ComVta_Dol.Caption = Format(CDbl(pnl_ComVta_Sol.Caption) / l_dbl_TipCam, "###,###,##0.00") & "  "
      pnl_ApoPro_Dol.Caption = Format(CDbl(pnl_ApoPro_Sol.Caption) / l_dbl_TipCam, "###,###,##0.00") & "  "
      pnl_MtoPre_Dol.Caption = Format(CDbl(pnl_MtoPre_Sol.Caption) / l_dbl_TipCam, "###,###,##0.00") & "  "
   Else
      pnl_ComVta_Sol.Caption = Format(CDbl(pnl_ValTot.Caption) * l_dbl_TipCam, "###,###,##0.00") & "  "
      pnl_ApoPro_Sol.Caption = Format(CDbl(pnl_CuoIni.Caption) * l_dbl_TipCam, "###,###,##0.00") & "  "
      pnl_MtoPre_Sol.Caption = Format(CDbl(pnl_TotPre.Caption) * l_dbl_TipCam, "###,###,##0.00") & "  "
      pnl_ComVta_Dol.Caption = Format(CDbl(pnl_ValTot.Caption), "###,###,##0.00") & "  "
      pnl_ApoPro_Dol.Caption = Format(CDbl(pnl_CuoIni.Caption), "###,###,##0.00") & "  "
      pnl_MtoPre_Dol.Caption = Format(CDbl(pnl_TotPre.Caption), "###,###,##0.00") & "  "
   End If
   
   If CDbl(pnl_ValTot.Caption) > 0 And CDbl(ipp_ApoPro.Text) > 0 Then
      If moddat_g_str_CodPrd = "023" Then
         lbl_PorIni.Caption = "(" & Format(CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption) * 100, "##0.00") & "%)"
      Else
         lbl_PorIni.Caption = "(" & Format((CDbl(ipp_ApoPro.Text) + CDbl(Me.ipp_MtoAFP.Value)) / CDbl(pnl_ValTot.Caption) * 100, "##0.00") & "%)"
      End If
   End If
End Sub

Private Sub fs_Limpia_CuoFin()
   pnl_CuoFin_MPr.Caption = "0.00 "
   pnl_CuoFin_Sol.Caption = "0.00 "
End Sub

Private Sub fs_Buscar_LisOcu()
   Call gs_LimpiaGrid(grd_LisOcu)
   moddat_g_int_NumObs = 0
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 31 "
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
      
      'Descripción Ocurrencia
      grd_LisOcu.Col = 2
      grd_LisOcu.Text = moddat_gf_Consulta_ParDes("004", Format(g_rst_Princi!SEGDET_CODOCU, "000000"))
      
      If g_rst_Princi!SEGDET_CODOCU = 21 Then
         If g_rst_Princi!SEGFECACT > 0 Then
            grd_LisOcu.Text = grd_LisOcu.Text & " (DESCARGO EFECTUADO - " & gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
            grd_LisOcu.Text = grd_LisOcu.Text & " / " & gf_FormatoHora(Format(g_rst_Princi!SEGHORACT, "000000")) & ")"
         Else
            moddat_g_int_NumObs = g_rst_Princi!SEGDET_NUMOBS
         End If
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
      
      'Fecha de Excepción
      grd_LisExc.Col = 0
      grd_LisExc.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Excepción
      grd_LisExc.Col = 1
      grd_LisExc.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Instancia
      grd_LisExc.Col = 2
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGEXC_CODINS))
      
      'Descripción Excepción
      grd_LisExc.Col = 3
      grd_LisExc.Text = Trim(g_rst_Princi!SEGEXC_DESCRI & "")
      
      'Tipo Autorización
      grd_LisExc.Col = 4
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("243", CStr(g_rst_Princi!SEGEXC_TIPAUT))
      
      'Motivo de Excepción
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
      
      'Descripción Condiciones
      grd_LisCon.Col = 1
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSCON & "")
      
      'Situación
      grd_LisCon.Col = 2
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("244", CStr(g_rst_Princi!SEGCON_SITUAC))
      
      If g_rst_Princi!SEGCON_SITUAC = 1 Then
         l_int_AprCon = 1
      End If
      
      'Descripción Levantamiento Condiciones
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

Private Sub fs_RegExc()
Dim r_int_NumExc     As Integer
Dim r_int_NivAut     As Integer

   If modgen_g_int_TipUsu = 20220 Then
      r_int_NivAut = 31
   Else
      r_int_NivAut = 13
   End If

   'Generando Número de Excepción
   r_int_NumExc = 0
   g_str_Parame = "SELECT COUNT(SEGEXC_NUMSOL) AS NUMREG FROM TRA_SEGEXC WHERE "
   g_str_Parame = g_str_Parame & "SEGEXC_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
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
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 31, 18, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Grabando en Detalle de Excepciones
   If Not moddat_gf_Inserta_SegExc(moddat_g_str_NumSol, 31, r_int_NumExc, UCase(moddat_g_str_Msje02), r_int_NivAut) Then
      Exit Sub
   End If
End Sub

Private Sub fs_EvaCre()
   l_dbl_CuoApr = 0
   l_dbl_MtoPre_Cal = 0
   l_int_PlaAno_Cal = 0
   l_int_TipSeg_Cal = 0
   l_int_PerGra_Cal = 0
   l_int_CuoDbl_Cal = 0

   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      l_dbl_CuoApr = g_rst_Princi!EVACRE_CUOSOL
      l_dbl_MtoPre_Cal = g_rst_Princi!EVACRE_MTOPRE_CAL
      l_int_PlaAno_Cal = g_rst_Princi!EVACRE_PLAANO_CAL
      l_int_TipSeg_Cal = g_rst_Princi!EVACRE_TIPSEG_CAL
      l_int_PerGra_Cal = g_rst_Princi!EVACRE_PERGRA_CAL
      l_int_CuoDbl_Cal = g_rst_Princi!EVACRE_CUODBL_CAL
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

'**************************************************************************************************
'* CONTROLES
'**************************************************************************************************
Private Sub txt_ObsSol_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Descar_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsExc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_LevCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
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
      pnl_DesOcu.Caption = "Día: " & r_str_FecOcu & " - " & r_str_HorOcu & " hrs. - " & r_str_DesOcu
   
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
'   Call grd_LisOcu_Click
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
      pnl_DesExc.Caption = "Día: " & r_str_FecExc & " - " & r_str_HorExc & " hrs. - " & r_str_InsExc
   
      grd_LisExc.Col = 3
      txt_ObsExc.Text = grd_LisExc.Text
      
      grd_LisExc.Col = 4
      pnl_TipAut.Caption = grd_LisExc.Text
      
      If LCase(Trim(r_str_InsExc)) = LCase("EVALUACION CREDITICIA") Then
         grd_LisExc.Col = 5
         pnl_motivo.Caption = IIf(grd_LisExc.Text = "0", " ", grd_LisExc.Text)
         pnl_motivo.Visible = True
         lbl_motivo.Visible = True
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
'   Call grd_LisExc_Click
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
'   Call grd_LisCon_Click
End Sub

Private Sub ipp_ComVta_Change()
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
   Call fs_Calcul
   Call fs_Limpia_CuoFin
End Sub

Private Sub ipp_ComVta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEst)
   End If
End Sub

Private Sub ipp_ComVta_LostFocus()

   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
   Call fs_Calcul
End Sub

Private Sub ipp_ValEst_Change()

   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
   Call fs_Calcul
   Call fs_Limpia_CuoFin
End Sub

Private Sub ipp_ValEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ApoPro)
   End If
End Sub

Private Sub ipp_ValEst_LostFocus()

   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
   Call fs_Calcul
End Sub

Private Sub ipp_ApoPro_Change()
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
   
   Call fs_Calcul
   Call fs_Limpia_CuoFin
End Sub

Private Sub ipp_ApoPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoAFP)
   End If
End Sub

Private Sub ipp_ApoPro_LostFocus()
   Call fs_Calcular_Prestamo
End Sub

Private Sub ipp_MtoAFP_Change()

   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
   Call fs_Calcul
   Call fs_Limpia_CuoFin
End Sub

Private Sub ipp_MtoAFP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CuoDbl)
   End If
End Sub

Private Sub cmb_CuoDbl_Click()
   Call fs_Limpia_CuoFin
End Sub

Private Sub cmb_CuoDbl_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaAno)
   End If
End Sub

Private Sub ipp_PlaAno_Change()
   Call fs_Limpia_CuoFin
End Sub

Private Sub ipp_PlaAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerGra)
   End If
End Sub

Private Sub ipp_PerGra_Change()
   Call fs_Limpia_CuoFin
End Sub

Private Sub ipp_PerGra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DiaPag)
   End If
End Sub

Private Sub cmb_DiaPag_Click()
   Call fs_Limpia_CuoFin
   Call gs_SetFocus(cmb_EmpSeg)
End Sub

Private Sub cmb_DiaPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_DiaPag_Click
   End If
End Sub

Private Sub cmb_EmpSeg_Click()
   Call fs_Limpia_CuoFin
   If cmb_EmpSeg.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_TipSeg(cmb_SegDes, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo)
      Screen.MousePointer = 0
      Call gs_SetFocus(cmb_SegDes)
   Else
      cmb_SegDes.Clear
   End If
End Sub

Private Sub cmb_EmpSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpSeg_Click
   End If
End Sub

Private Sub cmb_SegDes_Click()
   Call fs_Limpia_CuoFin
   Call gs_SetFocus(cmb_TasEsp)
End Sub

Private Sub cmb_SegDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SegDes_Click
   End If
End Sub

Private Sub cmb_TasEsp_Click()
   Call fs_Limpia_CuoFin
   Call gs_SetFocus(cmd_DatInm)
End Sub

Private Sub cmb_TasEsp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TasEsp_Click
   End If
End Sub

Private Sub fs_Calcular_GCierre()
   
   If chk_Gastos.Value = True Then
      pnl_ValGas.Caption = Format(CDbl(gf_Genera_Gastos_Cierre(moddat_g_str_CodPrd, l_str_CodPry, Format(moddat_g_str_CodMod, "000"), CDbl(ipp_ComVta.Value), CDbl(ipp_ValEst.Value), CDbl(pnl_MtoPre.Caption), l_dbl_GasTas, l_dbl_GasNot, l_dbl_BloReg, l_dbl_RegMin, l_dbl_RegHip, l_dbl_ImpITF)), "##,###,##0.00") & " "
      
      If l_dbl_GasTas = 0 Then
         MsgBox "No se asignaron valores para los Gastos de Tasación.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
      ElseIf l_dbl_GasNot = 0 Then
         MsgBox "No se asignaron valores para los Gastos Notariales.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
      ElseIf l_dbl_BloReg = 0 And Format(moddat_g_str_CodMod, "000") = "001" Then
         MsgBox "No se asignaron valores para los Gastos Registrales - Bloqueo Registral.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
      ElseIf l_dbl_RegMin = 0 Then
         MsgBox "No se asignaron valores para los Gastos Registrales - Minuta de Compra Venta.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
      ElseIf l_dbl_RegHip = 0 Then
         MsgBox "No se asignaron valores para los Gastos Registrales - Inscripción de Garantía.", vbExclamation, modgen_g_str_NomPlt
         chk_Gastos.Value = False
      End If
      
      If CDbl(pnl_ValGas.Caption) < 0 Then pnl_ValGas.Caption = "0.00 "
      If CDbl(pnl_ValGas.Caption) = 0 And CDbl(ipp_ValEst.Value = 0) Then pnl_ValGas.Caption = "0.00 "
   Else
      pnl_ValGas.Caption = "0.00 "
   End If
   pnl_TotPre.Caption = Format(CDbl(pnl_MtoPre.Caption) + CDbl(pnl_ValGas.Caption), "##,###,##0.00") & " "
End Sub

Private Sub fs_Calcular_Prestamo()
   'Valor inmueble
   pnl_ValTot.Caption = Format(CDbl(ipp_ComVta.Value) + CDbl(ipp_ValEst.Value), "##,###,##0.00") & " "
   If CDbl(pnl_ValTot.Caption) < 0 Then pnl_ValTot.Caption = "0.00 "
   
   'Inicial
   Call fs_Calcular_MtoSBMS
   Call fs_Bono_Verde(l_dbl_BmsTas)
   
   pnl_CuoIni.Caption = Format(CDbl(ipp_ApoPro.Value) + CDbl(pnl_FmvBbp.Caption) + CDbl(ipp_MtoAFP.Text) + CDbl(pnl_MtoBMS.Caption), "##,###,##0.00") & " "
   If CDbl(pnl_CuoIni.Caption) < 0 Then pnl_CuoIni.Caption = "0.00 "
   
   'Prestamo
   pnl_MtoPre.Caption = Format(CDbl(pnl_ValTot.Caption) - CDbl(pnl_CuoIni.Caption), "##,###,##0.00") & " "
   If CDbl(pnl_MtoPre.Caption) < 0 Then pnl_MtoPre.Caption = "0.00 "
   pnl_TotPre.Caption = Format(CDbl(pnl_MtoPre.Caption) + CDbl(pnl_ValGas.Caption), "##,###,##0.00") & " "
   DoEvents
End Sub

Private Sub fs_Calcular_MtoSBMS()
   l_dbl_MPSMS = Format(CDbl(pnl_ValTot.Caption) - CDbl(ipp_ApoPro.Text) - (CDbl(pnl_FmvBbp.Caption)) - CDbl(ipp_MtoAFP.Text), "###,###,##0.00") & " "
   If l_dbl_MPSMS < 0 Then l_dbl_MPSMS = 0
End Sub

Private Sub fs_Bono_Verde(ByVal p_ValAfe As Double)
   If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
      If p_ValAfe > 0 Then
         If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin Then
            pnl_MtoBMS.Caption = Format((CDbl(l_dbl_MPSMS) * modatecli_g_dbl_BMSTas) / (1 + modatecli_g_dbl_BMSTas), "###,###,##0.00") & "  "
         Else
            pnl_MtoBMS.Caption = Format((CDbl(l_dbl_MPSMS) * p_ValAfe) / (1 + p_ValAfe), "###,###,##0.00") & "  "
         End If
         If pnl_MtoBMS.Caption < 0 Then pnl_MtoBMS.Caption = "0.00  "
      End If
      
      pnl_MtoPre.Caption = Format(CDbl(l_dbl_MPSMS) - CDbl(pnl_MtoBMS.Caption), "###,###,##0.00") & "  "
      If pnl_MtoPre.Caption < 0 Then pnl_MtoPre.Caption = "0.00  "
   Else
      pnl_MtoBMS.Caption = "0.00  "
   End If
End Sub

Private Function fs_Actualiza_FinGCi(ByVal p_NumSol As String, p_GasCie As Double) As Integer
Dim r_rst_Genera     As ADODB.Recordset
Dim r_str_Parame     As String
Dim r_str_Operac     As String
Dim r_dbl_PorITF     As Double
   
   fs_Actualiza_FinGCi = False
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT SOLMAE_MTOGCI "
   r_str_Parame = r_str_Parame & "   FROM CRE_SOLMAE  "
   r_str_Parame = r_str_Parame & "  WHERE SOLMAE_NUMERO = '" & p_NumSol & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      'Con Financiamiento
      If CDbl(r_rst_Genera!SOLMAE_MTOGCI) > 0 Then
         
         If p_GasCie <> 0 Then
         
            If fs_ValPag_GasCie(p_NumSol, 1) = False Then
               'Actualiza los gastos de Cierre
               If l_dbl_GasTas > 0 Then
                  If Not fs_Actualiza_GasCie(p_NumSol, 11, l_dbl_GasTas) Then
                     fs_Actualiza_FinGCi = False
                  End If
               End If
               If l_dbl_GasNot > 0 Then
                  If Not fs_Actualiza_GasCie(p_NumSol, 12, l_dbl_GasNot) Then
                     fs_Actualiza_FinGCi = False
                  End If
               End If
               If l_dbl_BloReg > 0 Then
                  If Not fs_Actualiza_GasCie(p_NumSol, 13, l_dbl_BloReg) Then
                     fs_Actualiza_FinGCi = False
                  End If
               End If
               If l_dbl_RegMin > 0 Then
                  If Not fs_Actualiza_GasCie(p_NumSol, 16, l_dbl_RegMin) Then
                     fs_Actualiza_FinGCi = False
                  End If
               End If
               If l_dbl_RegHip > 0 Then
                  If Not fs_Actualiza_GasCie(p_NumSol, 17, l_dbl_RegHip) Then
                     fs_Actualiza_FinGCi = False
                  End If
               End If
               If l_dbl_ImpITF > 0 Then
                  If Not fs_Actualiza_GasCie(p_NumSol, 18, l_dbl_ImpITF) Then
                     fs_Actualiza_FinGCi = False
                  End If
               End If
            End If

            fs_Actualiza_FinGCi = True
            
         ElseIf p_GasCie = 0 Then
         
            If fs_ValPag_GasCie(p_NumSol, 1) = True Then             'Si pagó Tasación
               fs_Actualiza_FinGCi = True
            Else                                                     'No pagó Tasación
               If fs_Elimina_GasCie(p_NumSol) Then
                  
                  'GASTOS DE TASACION
                  If l_dbl_GasTas > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 111, l_dbl_GasTas) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  'GASTOS NOTARIALES
                  If l_dbl_GasNot > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 121, l_dbl_GasNot) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  'ITF
                  If l_dbl_ImpITF > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 131, l_dbl_ImpITF) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  'GASTOS REGISTRALES - BLOQUEO REGISTRAL (Si es Bien Terminado)
                  If l_dbl_BloReg > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 161, l_dbl_BloReg) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  'GASTOS REGISTRALES - MINUTA DE COMPRA VENTA
                  If l_dbl_RegMin > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 171, l_dbl_RegMin) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  'GASTOS REGISTRALES - INCRIPCION DE GARANTIA
                  If l_dbl_RegHip > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 181, l_dbl_RegHip) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
               End If
               
               'Actualiza en CRE_SOLMAE
               If Not fs_Actualiza_Solmae(p_NumSol, CDbl(pnl_TotPre.Caption), CDbl(p_GasCie)) Then
                  fs_Actualiza_FinGCi = False
               End If
            End If
            
            fs_Actualiza_FinGCi = True
         End If
         
      'Sin Financiamiento
      Else
         If p_GasCie = 0 Then
            fs_Actualiza_FinGCi = True
         Else
            'If fs_ValPag_GasCie(p_NumSol, 2) Then     'Si ya están pagados
            If fs_ValPag_GasCie(p_NumSol, 3) Then     'Si ya están pagados
               fs_Actualiza_FinGCi = True
            Else                                      'Si no han sido pagados
               If fs_Elimina_GasCie(p_NumSol) Then
                  
                  'GASTOS DE TASACION
                  If l_dbl_GasTas > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 111, l_dbl_GasTas) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  
                  'GASTOS NOTARIALES
                  If l_dbl_GasNot > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 121, l_dbl_GasNot) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  
                  'ITF
                  If l_dbl_ImpITF > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 131, l_dbl_ImpITF) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  
                  'Si es Bien Terminado
                  'GASTOS REGISTRALES - BLOQUEO REGISTRAL
                  If l_dbl_BloReg > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 161, l_dbl_BloReg) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  
                  'GASTOS REGISTRALES - MINUTA DE COMPRA VENTA
                  If l_dbl_RegMin > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 171, l_dbl_RegMin) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  
                  'GASTOS REGISTRALES - INCRIPCION DE GARANTIA
                  If l_dbl_RegHip > 0 Then
                     If Not fs_Asigna_GasAdm(p_NumSol, 181, l_dbl_RegHip) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  
                  'Pago de Gastos de Cierre, Menos Tasación
                  r_str_Operac = "99999"
         
                  'Obteniendo ITF
                  r_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
                  
                  'GASTOS NOTARIALES
                  If Not opecaj_gf_Pago_GasAdm(p_NumSol, Left(121, 2), moddat_g_int_TipMon, l_dbl_GasNot, r_dbl_PorITF, Format(CDate(Now), "yyyymmdd"), r_str_Operac) Then
                     fs_Actualiza_FinGCi = False
                  End If
                  
                  'ITF
                  If Not opecaj_gf_Pago_GasAdm(p_NumSol, Left(131, 2), moddat_g_int_TipMon, l_dbl_ImpITF, r_dbl_PorITF, Format(CDate(Now), "yyyymmdd"), r_str_Operac) Then
                     fs_Actualiza_FinGCi = False
                  End If
                  
                  'Si es Bien Terminado
                  'GASTOS REGISTRALES - BLOQUEO REGISTRAL
                  If l_dbl_BloReg > 0 Then
                     If Not opecaj_gf_Pago_GasAdm(p_NumSol, Left(161, 2), moddat_g_int_TipMon, l_dbl_BloReg, r_dbl_PorITF, Format(CDate(Now), "yyyymmdd"), r_str_Operac) Then
                        fs_Actualiza_FinGCi = False
                     End If
                  End If
                  
                  'GASTOS REGISTRALES - MINUTA DE COMPRA VENTA
                  If Not opecaj_gf_Pago_GasAdm(p_NumSol, Left(171, 2), moddat_g_int_TipMon, l_dbl_RegMin, r_dbl_PorITF, Format(CDate(Now), "yyyymmdd"), r_str_Operac) Then
                     fs_Actualiza_FinGCi = False
                  End If
                  
                  'GASTOS REGISTRALES - INCRIPCION DE GARANTIA
                  If Not opecaj_gf_Pago_GasAdm(p_NumSol, Left(181, 2), moddat_g_int_TipMon, l_dbl_RegHip, r_dbl_PorITF, Format(CDate(Now), "yyyymmdd"), r_str_Operac) Then
                     fs_Actualiza_FinGCi = False
                  End If
                  
                  'Actualizando en Seguimiento de Tasacion Pago de Gastos Administrativos
                  If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 31, 24, 0, "", 0, 0) Then
                     fs_Actualiza_FinGCi = False
                  End If
                  
                  'Actualiza en CRE_SOLMAE
                  If Not fs_Actualiza_Solmae(p_NumSol, CDbl(pnl_TotPre.Caption), CDbl(p_GasCie)) Then
                     fs_Actualiza_FinGCi = False
                  End If
                  
                  fs_Actualiza_FinGCi = True
               End If
            End If
         End If
      End If
      
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
   End If
End Function

Private Function fs_Valida_FinGCi(ByVal p_NumSol As String, p_GasCie As Double) As Integer
Dim r_rst_Genera     As ADODB.Recordset
Dim r_str_Parame     As String
Dim r_str_Operac     As String
Dim r_dbl_PorITF     As Double
   
   fs_Valida_FinGCi = False
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT SOLMAE_MTOGCI "
   r_str_Parame = r_str_Parame & "   FROM CRE_SOLMAE  "
   r_str_Parame = r_str_Parame & "  WHERE SOLMAE_NUMERO = '" & p_NumSol & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      'Con Financiamiento
      If CDbl(r_rst_Genera!SOLMAE_MTOGCI) > 0 Then
         If p_GasCie <> 0 Then
            If CDbl(p_GasCie) = CDbl(r_rst_Genera!SOLMAE_MTOGCI) Then
               fs_Valida_FinGCi = True
            Else
               If fs_ValPag_GasCie(p_NumSol, 1) = True Then
                  MsgBox "El Gasto de Tasación ya ha sido pagado.", vbExclamation, modgen_g_str_NomPlt
               Else
                  fs_Valida_FinGCi = True
               End If
            End If
         ElseIf p_GasCie = 0 Then
            If fs_ValPag_GasCie(p_NumSol, 1) = True Then             'Si pagó Tasación
               MsgBox "El Gasto de Tasación ya ha sido pagado.", vbExclamation, modgen_g_str_NomPlt
            Else                                                     'No pagó Tasación
               fs_Valida_FinGCi = True
            End If
         End If
      'Sin Financiamiento
      Else
         If p_GasCie = 0 Then
            fs_Valida_FinGCi = True
         Else
            'If fs_ValPag_GasCie(p_NumSol, 2) Then     'Si ya están pagados
            If fs_ValPag_GasCie(p_NumSol, 3) Then     'Si ya están pagados
               MsgBox "Los Gastos de Cierre ya han sido pagados.", vbExclamation, modgen_g_str_NomPlt
            Else                                      'Si no han sido pagados
               fs_Valida_FinGCi = True
            End If
         End If
      End If
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
   End If
End Function

Private Function fs_ValPag_GasCie(ByVal p_NumSol As String, ByVal p_Opcion As Integer) As Boolean
Dim r_str_Parame  As String
   
   fs_ValPag_GasCie = False
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT * FROM TRA_GASADM "
   r_str_Parame = r_str_Parame & "  WHERE GASADM_NUMSOL = '" & p_NumSol & "' "
   If p_Opcion = 1 Then
      r_str_Parame = r_str_Parame & "    AND GASADM_CODGAS = 11 "
   ElseIf p_Opcion = 2 Then
      r_str_Parame = r_str_Parame & "    AND GASADM_CODGAS IN (11, 12, 13, 16, 17, 18) "
   Else
      r_str_Parame = r_str_Parame & "    AND GASADM_CODGAS IN (12, 13, 16, 17, 18) "
   End If
      
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      
      Do While Not g_rst_GenAux.EOF
         If g_rst_GenAux!GASADM_SITUAC = 1 Then
            fs_ValPag_GasCie = True               'PAGADO
            Exit Function
         Else
            fs_ValPag_GasCie = False              'NO PAGADO
         End If
         g_rst_GenAux.MoveNext
      Loop
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Function

Private Function fs_Actualiza_GasCie(ByVal p_NumSol As String, p_CodGas As Integer, p_Importe As Double) As Integer
Dim r_str_Parame  As String
   
   fs_Actualiza_GasCie = False
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " UPDATE TRA_GASADM SET  "
   r_str_Parame = r_str_Parame & "        GASADM_IMPORT = " & p_Importe & " "
   r_str_Parame = r_str_Parame & "  WHERE GASADM_NUMSOL = '" & p_NumSol & "' "
   r_str_Parame = r_str_Parame & "    AND GASADM_CODGAS = " & p_CodGas & " "
      
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 2) Then
       Exit Function
   End If
   
'   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   fs_Actualiza_GasCie = True
End Function

Private Function fs_Actualiza_Solmae(ByVal p_NumSol As String, p_TotPre As Double, ByVal p_GasCie As Double) As Integer
Dim r_str_Parame  As String
   
   fs_Actualiza_Solmae = False
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " UPDATE CRE_SOLMAE SET "
   r_str_Parame = r_str_Parame & "        SOLMAE_MTOGCI = " & CDbl(p_GasCie) & ", "
   r_str_Parame = r_str_Parame & "        SOLMAE_MTOPRE_SOL = " & CDbl(p_TotPre) & ", "
   r_str_Parame = r_str_Parame & "        SOLMAE_MTOPRE_MPR = " & CDbl(p_TotPre) & "  "
   r_str_Parame = r_str_Parame & "  WHERE SOLMAE_NUMERO = '" & p_NumSol & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 2) Then
       Exit Function
   End If
   
'   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   fs_Actualiza_Solmae = True
End Function

Private Function fs_Elimina_GasCie(ByVal p_NumSol As String) As Integer
Dim r_str_Parame  As String
   
   fs_Elimina_GasCie = False
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " DELETE FROM TRA_GASADM "
   r_str_Parame = r_str_Parame & "  WHERE GASADM_NUMSOL = '" & p_NumSol & "' "
   r_str_Parame = r_str_Parame & "    AND GASADM_CODGAS IN (11,12,13,16,17,18) "
      
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 2) Then
       Exit Function
   End If
   
'   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   fs_Elimina_GasCie = True
End Function

Private Function fs_Asigna_GasAdm(p_NumSol As String, p_Codigo As Integer, p_Import As Double) As Integer
   fs_Asigna_GasAdm = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_GASADM_ASIGNA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & Left(p_Codigo, 2) & ", "
      g_str_Parame = g_str_Parame & Right(p_Codigo, 1) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(p_Import)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   fs_Asigna_GasAdm = True
End Function
