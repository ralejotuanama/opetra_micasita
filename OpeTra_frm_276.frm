VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Tra_EvaTas_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9720
   ClientLeft      =   570
   ClientTop       =   1950
   ClientWidth     =   11595
   Icon            =   "OpeTra_frm_276.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   18018
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
         Height          =   3525
         Left            =   30
         TabIndex        =   1
         Top             =   4080
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6218
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
            Height          =   3405
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   6006
            _Version        =   393216
            Style           =   1
            Tab             =   2
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento de Instancias"
            TabPicture(0)   =   "OpeTra_frm_276.frx":000C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Label11"
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(2)=   "Label7"
            Tab(0).Control(3)=   "pnl_DesOcu"
            Tab(0).Control(4)=   "SSPanel5"
            Tab(0).Control(5)=   "SSPanel14"
            Tab(0).Control(6)=   "SSPanel13"
            Tab(0).Control(7)=   "grd_LisOcu"
            Tab(0).Control(8)=   "SSPanel10"
            Tab(0).Control(9)=   "txt_Descar"
            Tab(0).Control(10)=   "txt_Observ"
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "OpeTra_frm_276.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txt_ObsExc"
            Tab(1).Control(1)=   "grd_LisExc"
            Tab(1).Control(2)=   "SSPanel9"
            Tab(1).Control(3)=   "SSPanel11"
            Tab(1).Control(4)=   "SSPanel12"
            Tab(1).Control(5)=   "SSPanel15"
            Tab(1).Control(6)=   "SSPanel16"
            Tab(1).Control(7)=   "pnl_DesExc"
            Tab(1).Control(8)=   "pnl_TipAut"
            Tab(1).Control(9)=   "pnl_motivo"
            Tab(1).Control(10)=   "lbl_motivo"
            Tab(1).Control(11)=   "Label6"
            Tab(1).Control(12)=   "Label3"
            Tab(1).Control(13)=   "Label4"
            Tab(1).ControlCount=   14
            TabCaption(2)   =   "Aprobación Condicionada"
            TabPicture(2)   =   "OpeTra_frm_276.frx":0044
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Label12"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "Label14"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "Label15"
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
            Tab(2).Control(9)=   "txt_LevCon"
            Tab(2).Control(9).Enabled=   0   'False
            Tab(2).Control(10)=   "txt_ObsCon"
            Tab(2).Control(10).Enabled=   0   'False
            Tab(2).ControlCount=   11
            Begin VB.TextBox txt_ObsCon 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Text            =   "OpeTra_frm_276.frx":0060
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Text            =   "OpeTra_frm_276.frx":0064
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   975
               Left            =   -73770
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Text            =   "OpeTra_frm_276.frx":0068
               Top             =   1980
               Width           =   10065
            End
            Begin VB.TextBox txt_Observ 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Text            =   "OpeTra_frm_276.frx":006C
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_Descar 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "OpeTra_frm_276.frx":0070
               Top             =   2640
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
               Width           =   10035
               _Version        =   65536
               _ExtentX        =   17701
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
               TabIndex        =   21
               Top             =   2970
               Width           =   3765
               _Version        =   65536
               _ExtentX        =   6641
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
               Left            =   2790
               TabIndex        =   26
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
               Left            =   1320
               TabIndex        =   27
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
               Left            =   -68970
               TabIndex        =   68
               Top             =   2970
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
                  TabIndex        =   69
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
               Left            =   -69720
               TabIndex        =   70
               Top             =   3030
               Width           =   645
            End
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobación:"
               Height          =   495
               Left            =   60
               TabIndex        =   36
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   60
               TabIndex        =   35
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label12 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   615
               Left            =   60
               TabIndex        =   34
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   33
               Top             =   2970
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Excepción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   32
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripción:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   31
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observación:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   30
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   29
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   28
               Top             =   2640
               Width           =   1035
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   6810
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   720
            TabIndex        =   57
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito Hipotecario"
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
            Left            =   720
            TabIndex        =   58
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Tasación del Inmueble"
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
            Left            =   90
            Picture         =   "OpeTra_frm_276.frx":0074
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
         Begin VB.CommandButton cmd_DatInm 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_276.frx":037E
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Modificación de Datos del Inmueble"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueObs 
            Height          =   585
            Left            =   1800
            Picture         =   "OpeTra_frm_276.frx":0C48
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Registro de Observación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Excepc 
            Height          =   585
            Left            =   2400
            Picture         =   "OpeTra_frm_276.frx":108A
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Registro de Excepción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_AprCon 
            Height          =   585
            Left            =   3000
            Picture         =   "OpeTra_frm_276.frx":1394
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Aprobación con Condición"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   585
            Left            =   4200
            Picture         =   "OpeTra_frm_276.frx":169E
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   585
            Left            =   3600
            Picture         =   "OpeTra_frm_276.frx":1AE0
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_276.frx":1DEA
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Evalua 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_276.frx":222C
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Registrar Informe de Tasación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_OrdTas 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_276.frx":2AF6
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Generar Orden de Tasación"
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
         Begin VB.Label Label2 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   8070
            TabIndex        =   54
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   53
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   52
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   2055
         Left            =   0
         TabIndex        =   55
         Top             =   7650
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3625
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisEva 
            Height          =   1965
            Left            =   30
            TabIndex        =   56
            Top             =   60
            Width           =   11445
            _ExtentX        =   20188
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   1785
         Left            =   30
         TabIndex        =   59
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
            TabIndex        =   60
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
            TabPicture(0)   =   "OpeTra_frm_276.frx":2E00
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "OpeTra_frm_276.frx":2E1C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Apoderado"
            TabPicture(2)   =   "OpeTra_frm_276.frx":2E38
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Inmueble"
            TabPicture(3)   =   "OpeTra_frm_276.frx":2E54
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos Crédito"
            TabPicture(4)   =   "OpeTra_frm_276.frx":2E70
            Tab(4).ControlEnabled=   -1  'True
            Tab(4).Control(0)=   "Label5"
            Tab(4).Control(0).Enabled=   0   'False
            Tab(4).Control(1)=   "grd_Listad(4)"
            Tab(4).Control(1).Enabled=   0   'False
            Tab(4).Control(2)=   "txt_ObsSol"
            Tab(4).Control(2).Enabled=   0   'False
            Tab(4).ControlCount=   3
            Begin VB.TextBox txt_ObsSol 
               Height          =   405
               Left            =   1290
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   61
               Text            =   "OpeTra_frm_276.frx":2E8C
               Top             =   1200
               Width           =   10005
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1245
               Index           =   0
               Left            =   -74940
               TabIndex        =   62
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
               TabIndex        =   63
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
               Index           =   3
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
               Height          =   825
               Index           =   4
               Left            =   60
               TabIndex        =   66
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
            Begin VB.Label Label5 
               Caption         =   "Observaciones:"
               Height          =   300
               Left            =   90
               TabIndex        =   67
               Top             =   1200
               Width           =   1155
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_EvaTas_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_AprCon     As Integer
Dim l_dbl_MtoPre     As Double
Dim l_dbl_SumAse     As Double

Private Sub cmd_AprCon_Click()
Dim r_int_TipDoc     As Integer
Dim r_int_CodAct     As Integer
Dim r_int_Contad     As Integer
Dim r_int_FlgDoc     As Integer
Dim r_int_DiaTra     As Integer
Dim r_str_CodGrp     As String
Dim r_str_CodIte     As String

   If grd_LisEva.Rows = 0 Then
      MsgBox "No se ha registrado información de la Evaluación del Peritaje del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If moddat_gf_Valida_Observ(moddat_g_str_NumSol, 41) Then
      MsgBox "La solicitud presenta Observaciones pendientes de descargo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If pnl_AprCon.Visible Then
      MsgBox "La solicitud presenta Aprobación Condicionada pendiente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If modgen_g_int_TipUsu <> 18200 And modgen_g_int_TipUsu <> 18000 Then
      If l_dbl_MtoPre > l_dbl_SumAse Then
         MsgBox "Atención!!! La monto del préstamo no puede ser mayor que la suma asegurada.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   Else
      If l_dbl_MtoPre > l_dbl_SumAse Then
         MsgBox "Atención!!! La monto del préstamo no puede ser mayor que la suma asegurada.", vbExclamation, modgen_g_str_NomPlt
      End If
   End If
   
   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_str_Observ = ""
   moddat_g_int_FlgAct_1 = 1
   
   frm_RecSol_56.Show 1

   If moddat_g_int_FlgAct_1 = 1 Then
      Exit Sub
   End If
   
   'Creando Aprobación Condicionada
   If Not moddat_gf_Inserta_AprCon(moddat_g_str_NumSol, 41, moddat_g_str_Observ) Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 41)))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 41, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Verificar si aprobo Seguros
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 42"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   g_rst_Genera.MoveFirst

   If g_rst_Genera!SEGUIM_SITUAC <> 1 Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
      Unload Me
   Else
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      'Inserta Nueva Instancia de Evaluación
      If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 51) Then
         Exit Sub
      End If
         
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 51, 11, 0, "", 0, 0) Then
         Exit Sub
      End If
      
      'Actualizando en Tabla de Créditos
      If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, 51) Then
         Exit Sub
      End If
      
      'Enviando Correo Electrónico
      modgen_g_str_Mail_Asunto = "TASACION DEL INMUEBLE Y EVALUACION DE SEGUROS - APROBACION CONDICIONADA (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ & Chr(13)
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, True, True)
      
      MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
      Unload Me
   End If
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
      MsgBox "No se ha registrado información de la Evaluación del Peritaje del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If moddat_gf_Valida_Observ(moddat_g_str_NumSol, 41) Then
      MsgBox "La solicitud presenta Observaciones pendientes de descargo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If pnl_AprCon.Visible Then
      MsgBox "La solicitud presenta Aprobación Condicionada pendiente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If modgen_g_int_TipUsu <> 18200 And modgen_g_int_TipUsu <> 18000 And fs_Validar_TipInm = True Then
      If l_dbl_MtoPre > l_dbl_SumAse Then
         MsgBox "Atención!!! El monto del préstamo no puede ser mayor que la suma asegurada.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   Else
      If l_dbl_MtoPre > l_dbl_SumAse Then
         MsgBox "Atención!!! El monto del préstamo no puede ser mayor que la suma asegurada.", vbExclamation, modgen_g_str_NomPlt
      End If
   End If

   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 41)))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 41, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Verificar si aprobo Seguros
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 42"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   g_rst_Genera.MoveFirst

   If g_rst_Genera!SEGUIM_SITUAC <> 1 Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
      Unload Me
   Else
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      'Inserta Nueva Instancia de Evaluación
      If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 51) Then
         Exit Sub
      End If
         
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 51, 11, 0, "", 0, 0) Then
         Exit Sub
      End If
      
      'Actualizando en Tabla de Créditos
      If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, 51) Then
         Exit Sub
      End If
      
      'Enviando Correo Electrónico
      modgen_g_str_Mail_Asunto = "TASACION DEL INMUEBLE Y EVALUACION DE SEGUROS - APROBACION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, True, True)
      
      MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
      Unload Me
   End If
End Sub
Function fs_Validar_TipInm() As Boolean
   
   fs_Validar_TipInm = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT SOLINM_TIPINM, TRIM(PARDES_DESCRI) AS TIPO_INMUEBLE"
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLINM SL "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES ON PARDES_CODGRP = '217' AND PARDES_CODITE = SOLINM_TIPINM "
   g_str_Parame = g_str_Parame & "  WHERE SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CStr(g_rst_GenAux!TIPO_INMUEBLE) = "DEPARTAMENTO" Then
         fs_Validar_TipInm = True
      End If
   End If
End Function
Private Sub cmd_DatInm_Click()
   If moddat_g_int_InsAct >= 41 Then
      MsgBox "La información del Inmueble sólo puede ser modificada antes del envío a Tasación.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   moddat_g_int_FlgGrb = 2
   frm_Seg_SolHip_54.Show 1
End Sub

Private Sub cmd_Evalua_Click()
   moddat_g_int_FlgAct = 1
   'frm_Tra_EvaTas_04.Show 1
   frm_ModSol_06.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar_DatEva      'Buscando Información de Evaluación ya registrada
      Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
      Call fs_Buscar_LisExc      'Buscando Excepciones
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
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 18, 0, "", 0, 0) Then
         Exit Sub
      End If
      
      'Grabando en Detalle de Excepciones
      If Not moddat_gf_Inserta_SegExc(moddat_g_str_NumSol, 41, r_int_NumExc, moddat_g_str_Observ, moddat_g_int_TipAut) Then
         Exit Sub
      End If
      
      'Enviando Correo Electrónico
      modgen_g_str_Mail_Asunto = "TASACION DEL INMUEBLE - EXCEPCION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, True)
      
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
         'Generando Número de Observación
         r_int_NumObs = 0
            
         g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
         g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
         g_str_Parame = g_str_Parame & "SEGDET_CODINS = 41 AND "
         g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 21 "
         g_str_Parame = g_str_Parame & "ORDER BY SEGDET_NUMOBS DESC"
      
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
         If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 21, CStr(r_int_NumObs), moddat_g_str_Observ, 1, 0) Then
            Exit Sub
         End If
         
         'Actualizando en Instancia si es una Observación
         If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 41, 0, 3, 2) Then
            Exit Sub
         End If
         
         modgen_g_str_Mail_Asunto = "TASACION DEL INMUEBLE - OBSERVACION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      Else
         'Grabando en Detalle de Seguimiento
         If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 17, 0, moddat_g_str_Observ, 0, 0) Then
            Exit Sub
         End If
         
         modgen_g_str_Mail_Asunto = "TASACION DEL INMUEBLE - COMENTARIO (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      End If
      
      'Enviando Correo Electrónico
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, True)
      
      'Cargando Datos de Seguimiento
      Call fs_Buscar_LisOcu
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_OrdTas_Click()
   moddat_g_int_FlgAct_1 = 1

   frm_Tra_EvaTas_03.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Rechaz_Click()
Dim r_int_DiaTra     As Integer
Dim r_str_CodIns     As String
Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = 41
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 41)))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 41, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      'Actualizando Rechazo en Tabla de Créditos
      If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      modgen_g_str_Mail_Asunto = "TASACION DEL INMUEBLE - RECHAZO (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      
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
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, True)
   
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
      Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_arr_Mtz()   As moddat_g_tpo_DatCom
   
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   moddat_g_int_CodIns = 41
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   
   'Buscar Información de Solicitud de Crédito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""

   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Información del Cónyuge
   Call modmip_gs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(2))         'Buscar Información del Apoderado
   Call modmip_gs_DatInm(grd_Listad(3), False)                                            'Buscar Información del Inmueble
         
   'Buscar Información del Crédito
   Call modmip_gs_DatCre(grd_Listad(4), r_arr_Mtz)
   txt_ObsSol.Text = r_arr_Mtz(0).DatCom_Observ
   moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
   moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
   moddat_g_str_FecIng = r_arr_Mtz(0).DatCom_FecSol
   l_dbl_MtoPre = r_arr_Mtz(0).DatCom_MtoPre_Mpr
   
   'Call fs_DatCre
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
   
   'Cargando Datos de Evaluación Crediticia
   Call fs_Buscar_DatEva
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_int_Contad     As Integer

   l_dbl_MtoPre = 0
   l_dbl_SumAse = 0

   'Inicializando Grid de Cliente y de Cónyuge
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

   'Lista de Datos de Evaluación
   grd_LisEva.ColWidth(0) = 3300
   grd_LisEva.ColWidth(1) = 7940
   grd_LisEva.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEva.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Buscar_LisOcu()
Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisOcu)
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 41    "
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
         grd_LisExc.Col = 5
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
   
   Call grd_LisOcu_Click
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub fs_Buscar_DatEva()
   Call gs_LimpiaGrid(grd_LisEva)
   
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      l_dbl_SumAse = g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Empresa Peritaje"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("507", g_rst_Princi!EVATAS_CODEMP)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Nombre Perito"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = Trim(g_rst_Princi!EVATAS_NOMPER & "")
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Código REPEV SBS"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = Trim(g_rst_Princi!EVATAS_CODPER & "")
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Nro. de Informe"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = Trim(g_rst_Princi!EVATAS_NUMINF & "")
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Fecha Evaluación"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Año de Construcción"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = CStr(g_rst_Princi!EVATAS_ANOCON)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Nro. de Pisos"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = CStr(g_rst_Princi!EVATAS_NUMPIS)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Nro. de Sótanos"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = CStr(g_rst_Princi!EVATAS_NUMSOT)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Tipo de Inmueble"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("221", CStr(g_rst_Princi!EVATAS_TIPINM))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Uso de Inmueble"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("222", CStr(g_rst_Princi!EVATAS_USOINM))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Material de Construcción"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("223", CStr(g_rst_Princi!EVATAS_MATCON))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Tipo de Moneda"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!EVATAS_TIPMON))
      
      'Total
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Area Terreno (Total)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM + g_rst_Princi!EVATAS_ARETER_ES1 + g_rst_Princi!EVATAS_ARETER_ES2 + g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Area Construida (Total)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM + g_rst_Princi!EVATAS_ARECON_ES1 + g_rst_Princi!EVATAS_ARECON_ES2 + g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Suma Asegurada (Total)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Valor Comercial (Total)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM + g_rst_Princi!EVATAS_VALCOM_ES1 + g_rst_Princi!EVATAS_VALCOM_ES2 + g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Valor Realización (Total)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM + g_rst_Princi!EVATAS_VALREA_ES1 + g_rst_Princi!EVATAS_VALREA_ES2 + g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Valor Terreno (Total)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM + g_rst_Princi!EVATAS_VALTER_ES1 + g_rst_Princi!EVATAS_VALTER_ES2 + g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Valor Edificación (Total)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM + g_rst_Princi!EVATAS_VALEDI_ES1 + g_rst_Princi!EVATAS_VALEDI_ES2 + g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.Text = "Valor Areas Comunes (Total)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColNeg
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM + g_rst_Princi!EVATAS_VALACO_ES1 + g_rst_Princi!EVATAS_VALACO_ES2 + g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
   
      'Inmueble
      grd_LisEva.Rows = grd_LisEva.Rows + 2
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.Text = "Area Terreno (Inmueble)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM, 12, 2) & " m2"
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.Text = "Area Construida (Inmueble)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM, 12, 2) & " m2"
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.Text = "Suma Asegurada (Inmueble)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.Text = "Valor Comercial (Inmueble)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.Text = "Valor Realización (Inmueble)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.Text = "Valor Terreno (Inmueble)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.Text = "Valor Edificación (Inmueble)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.Text = "Valor Areas Comunes (Inmueble)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellForeColor = modgen_g_con_ColAzu
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM, 12, 2)
   
      'Estacionamiento 1
      If g_rst_Princi!EVATAS_FLGEST_ES1 = 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Area Terreno (Estac. 1)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES1, 12, 2) & " m2"
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Area Construida (Estac. 1)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES1, 12, 2) & " m2"
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Suma Asegurada (Estac. 1)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES1, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Valor Comercial (Estac. 1)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES1, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Valor Realización (Estac. 1)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES1, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Valor Terreno (Estac. 1)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES1, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Valor Edificación (Estac. 1)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES1, 12, 2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Valor Areas Comunes (Estac. 1)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES1, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_ES2 = 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.Text = "Area Terreno (Estac. 2)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES2, 12, 2) & " m2"
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.Text = "Area Construida (Estac. 2)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES2, 12, 2) & " m2"
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.Text = "Suma Asegurada (Estac. 2)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES2, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.Text = "Valor Comercial (Estac. 2)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES2, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.Text = "Valor Realización (Estac. 2)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES2, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.Text = "Valor Terreno (Estac. 2)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES2, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.Text = "Valor Edificación (Estac. 2)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES2, 12, 2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.Text = "Valor Areas Comunes (Estac. 2)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColNeg
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES2, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_DEP = 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Area Terreno (Depósito)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Area Construida (Depósito)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Suma Asegurada (Depósito)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Valor Comercial (Depósito)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Valor Realización (Depósito)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Valor Terreno (Depósito)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Valor Edificación (Depósito)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.Text = "Valor Areas Comunes (Depósito)"
         
         grd_LisEva.Col = 1
         grd_LisEva.CellForeColor = modgen_g_con_ColAzu
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
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

Private Sub txt_ObsExc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

