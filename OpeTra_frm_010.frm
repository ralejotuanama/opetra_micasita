VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_Des_CreHip_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   10665
   ClientLeft      =   1410
   ClientTop       =   375
   ClientWidth     =   12855
   Icon            =   "OpeTra_frm_010.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10665
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10665
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12855
      _Version        =   65536
      _ExtentX        =   22675
      _ExtentY        =   18812
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   1905
         Left            =   30
         TabIndex        =   14
         Top             =   7890
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   3360
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
         Begin TabDlg.SSTab tab_Cronog 
            Height          =   1815
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   12585
            _ExtentX        =   22199
            _ExtentY        =   3201
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Cliente - Tramo No Concesional"
            TabPicture(0)   =   "OpeTra_frm_010.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label3"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "pnl_CliNCo_TotCuo"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "pnl_CliNCo_OtrCar"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel62"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel61"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel59"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel14"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel13"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel12"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel11"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "SSPanel10"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "SSPanel9"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "grd_CliNCo_Listad"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_CliNCo_Intere"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_CliNCo_SegPre"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_CliNCo_SegViv"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_CliNCo_Capita"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).ControlCount=   17
            TabCaption(1)   =   "Cliente - Tramo Concesional"
            TabPicture(1)   =   "OpeTra_frm_010.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label13"
            Tab(1).Control(1)=   "pnl_CliCon_Capita"
            Tab(1).Control(2)=   "pnl_CliCon_Intere"
            Tab(1).Control(3)=   "pnl_CliCon_TotCuo"
            Tab(1).Control(4)=   "SSPanel25"
            Tab(1).Control(5)=   "SSPanel24"
            Tab(1).Control(6)=   "SSPanel22"
            Tab(1).Control(7)=   "SSPanel21"
            Tab(1).Control(8)=   "SSPanel19"
            Tab(1).Control(9)=   "SSPanel23"
            Tab(1).Control(10)=   "grd_CliCon_Listad"
            Tab(1).ControlCount=   11
            TabCaption(2)   =   "Cofide - Tramo No Concesional"
            TabPicture(2)   =   "OpeTra_frm_010.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label1"
            Tab(2).Control(1)=   "pnl_CofNCo_TotCuo"
            Tab(2).Control(2)=   "SSPanel37"
            Tab(2).Control(3)=   "pnl_CofNCo_Capita"
            Tab(2).Control(4)=   "pnl_CofNCo_Intere"
            Tab(2).Control(5)=   "pnl_CofNCo_Comisi"
            Tab(2).Control(6)=   "SSPanel32"
            Tab(2).Control(7)=   "SSPanel31"
            Tab(2).Control(8)=   "SSPanel29"
            Tab(2).Control(9)=   "SSPanel28"
            Tab(2).Control(10)=   "SSPanel27"
            Tab(2).Control(11)=   "SSPanel26"
            Tab(2).Control(12)=   "grd_CofNCo_Listad"
            Tab(2).ControlCount=   13
            TabCaption(3)   =   "Cofide - Tramo Concesional"
            TabPicture(3)   =   "OpeTra_frm_010.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Label2"
            Tab(3).Control(1)=   "pnl_CofCon_TotCuo"
            Tab(3).Control(2)=   "SSPanel57"
            Tab(3).Control(3)=   "pnl_CofCon_Capita"
            Tab(3).Control(4)=   "pnl_CofCon_Intere"
            Tab(3).Control(5)=   "pnl_CofCon_Comisi"
            Tab(3).Control(6)=   "SSPanel53"
            Tab(3).Control(7)=   "SSPanel52"
            Tab(3).Control(8)=   "SSPanel51"
            Tab(3).Control(9)=   "SSPanel50"
            Tab(3).Control(10)=   "SSPanel48"
            Tab(3).Control(11)=   "SSPanel40"
            Tab(3).Control(12)=   "grd_CofCon_Listad"
            Tab(3).ControlCount=   13
            Begin Threed.SSPanel pnl_CliNCo_Capita 
               Height          =   285
               Left            =   2550
               TabIndex        =   398
               Top             =   1470
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliNCo_SegViv 
               Height          =   285
               Left            =   6690
               TabIndex        =   401
               Top             =   1470
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliNCo_SegPre 
               Height          =   285
               Left            =   5310
               TabIndex        =   396
               Top             =   1470
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliNCo_Intere 
               Height          =   285
               Left            =   3930
               TabIndex        =   397
               Top             =   1470
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin MSFlexGridLib.MSFlexGrid grd_CliNCo_Listad 
               Height          =   795
               Left            =   30
               TabIndex        =   9
               Top             =   660
               Width           =   12465
               _ExtentX        =   21987
               _ExtentY        =   1402
               _Version        =   393216
               Rows            =   21
               Cols            =   9
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_CofCon_Listad 
               Height          =   795
               Left            =   -74970
               TabIndex        =   11
               Top             =   660
               Width           =   12465
               _ExtentX        =   21987
               _ExtentY        =   1402
               _Version        =   393216
               Rows            =   21
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_CofNCo_Listad 
               Height          =   795
               Left            =   -74970
               TabIndex        =   10
               Top             =   660
               Width           =   12465
               _ExtentX        =   21987
               _ExtentY        =   1402
               _Version        =   393216
               Rows            =   21
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_CliCon_Listad 
               Height          =   795
               Left            =   -74970
               TabIndex        =   8
               Top             =   660
               Width           =   12465
               _ExtentX        =   21987
               _ExtentY        =   1402
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
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -69870
               TabIndex        =   359
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   -74940
               TabIndex        =   356
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -73770
               TabIndex        =   357
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel22 
               Height          =   285
               Left            =   -72210
               TabIndex        =   358
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel24 
               Height          =   285
               Left            =   -67530
               TabIndex        =   360
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel25 
               Height          =   285
               Left            =   -65190
               TabIndex        =   361
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_CliCon_TotCuo 
               Height          =   285
               Left            =   -67530
               TabIndex        =   362
               Top             =   1470
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliCon_Intere 
               Height          =   285
               Left            =   -69870
               TabIndex        =   363
               Top             =   1470
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliCon_Capita 
               Height          =   285
               Left            =   -72210
               TabIndex        =   364
               Top             =   1470
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel26 
               Height          =   285
               Left            =   -74940
               TabIndex        =   366
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel27 
               Height          =   285
               Left            =   -73770
               TabIndex        =   367
               Top             =   360
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel28 
               Height          =   285
               Left            =   -71970
               TabIndex        =   368
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel29 
               Height          =   285
               Left            =   -70140
               TabIndex        =   369
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel31 
               Height          =   285
               Left            =   -66480
               TabIndex        =   370
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel32 
               Height          =   285
               Left            =   -64650
               TabIndex        =   371
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_CofNCo_Comisi 
               Height          =   285
               Left            =   -68310
               TabIndex        =   372
               Top             =   1470
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CofNCo_Intere 
               Height          =   285
               Left            =   -70140
               TabIndex        =   373
               Top             =   1470
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CofNCo_Capita 
               Height          =   285
               Left            =   -71970
               TabIndex        =   374
               Top             =   1470
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel37 
               Height          =   285
               Left            =   -68310
               TabIndex        =   376
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_CofNCo_TotCuo 
               Height          =   285
               Left            =   -66480
               TabIndex        =   377
               Top             =   1470
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel40 
               Height          =   285
               Left            =   -74940
               TabIndex        =   378
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel48 
               Height          =   285
               Left            =   -73770
               TabIndex        =   379
               Top             =   360
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel50 
               Height          =   285
               Left            =   -71970
               TabIndex        =   380
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel51 
               Height          =   285
               Left            =   -70140
               TabIndex        =   381
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel52 
               Height          =   285
               Left            =   -66480
               TabIndex        =   382
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel53 
               Height          =   285
               Left            =   -64650
               TabIndex        =   383
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_CofCon_Comisi 
               Height          =   285
               Left            =   -68310
               TabIndex        =   384
               Top             =   1470
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CofCon_Intere 
               Height          =   285
               Left            =   -70140
               TabIndex        =   385
               Top             =   1470
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CofCon_Capita 
               Height          =   285
               Left            =   -71970
               TabIndex        =   386
               Top             =   1470
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel57 
               Height          =   285
               Left            =   -68310
               TabIndex        =   388
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_CofCon_TotCuo 
               Height          =   285
               Left            =   -66480
               TabIndex        =   389
               Top             =   1470
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   60
               TabIndex        =   390
               Top             =   360
               Width           =   1125
               _Version        =   65536
               _ExtentX        =   1984
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   1170
               TabIndex        =   391
               Top             =   360
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   2550
               TabIndex        =   392
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   3930
               TabIndex        =   393
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel13 
               Height          =   285
               Left            =   9450
               TabIndex        =   394
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
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
               Left            =   10830
               TabIndex        =   395
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel59 
               Height          =   285
               Left            =   5310
               TabIndex        =   400
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel61 
               Height          =   285
               Left            =   6690
               TabIndex        =   402
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel62 
               Height          =   285
               Left            =   8070
               TabIndex        =   403
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Otros Cargos"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_CliNCo_OtrCar 
               Height          =   285
               Left            =   8070
               TabIndex        =   404
               Top             =   1470
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CliNCo_TotCuo 
               Height          =   285
               Left            =   9450
               TabIndex        =   405
               Top             =   1470
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin VB.Label Label3 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   1470
               TabIndex        =   399
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label2 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   387
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   375
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label13 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -73230
               TabIndex        =   365
               Top             =   1470
               Width           =   945
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            Height          =   495
            Left            =   630
            TabIndex        =   16
            Top             =   60
            Width           =   4905
            _Version        =   65536
            _ExtentX        =   8652
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Generación de Operación"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   405
            Left            =   4920
            TabIndex        =   17
            Top             =   150
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "DNI - 07521154 / IKEHARA PUNK MIGUEL ANGEL "
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
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_010.frx":007C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel36 
         Height          =   765
         Left            =   30
         TabIndex        =   18
         Top             =   9840
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   12000
            Picture         =   "OpeTra_frm_010.frx":0386
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   5115
         Left            =   30
         TabIndex        =   19
         Top             =   2730
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         BevelOuter      =   1
         Begin TabDlg.SSTab tab_Princi 
            Height          =   4995
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   8811
            _Version        =   393216
            Style           =   1
            Tabs            =   9
            TabsPerRow      =   9
            TabHeight       =   520
            TabCaption(0)   =   "Datos Cliente"
            TabPicture(0)   =   "OpeTra_frm_010.frx":07C8
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lbl_NomGlo(12)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lbl_NomGlo(13)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lbl_NomGlo(14)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "lbl_NomGlo(15)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lbl_NomGlo(16)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lbl_NomGlo(17)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "lbl_NomGlo(18)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "lbl_NomGlo(100)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "lbl_NomGlo(109)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "lbl_NomGlo(110)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "lbl_NomGlo(111)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "SSPanel30"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "pnl_Tit_RegCyg"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_Tit_DirEle"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_Tit_Direcc"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_Tit_Telefo"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_Tit_Celula"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "pnl_Tit_LugNac"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "pnl_Tit_Profes"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "pnl_Tit_NivEst"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "pnl_Tit_EstCiv"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "pnl_Tit_Paises"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "pnl_Tit_FecNac"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "tab_DatCli"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).ControlCount=   24
            TabCaption(1)   =   "Datos Cónyuge"
            TabPicture(1)   =   "OpeTra_frm_010.frx":07E4
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbl_NomGlo(34)"
            Tab(1).Control(1)=   "lbl_NomGlo(28)"
            Tab(1).Control(2)=   "lbl_NomGlo(35)"
            Tab(1).Control(3)=   "lbl_NomGlo(32)"
            Tab(1).Control(4)=   "lbl_NomGlo(31)"
            Tab(1).Control(5)=   "lbl_NomGlo(30)"
            Tab(1).Control(6)=   "lbl_NomGlo(29)"
            Tab(1).Control(7)=   "lbl_NomGlo(27)"
            Tab(1).Control(8)=   "lbl_NomGlo(26)"
            Tab(1).Control(9)=   "pnl_Cyg_ApeNom"
            Tab(1).Control(10)=   "pnl_Cyg_DocIde"
            Tab(1).Control(11)=   "SSPanel18"
            Tab(1).Control(12)=   "pnl_Cyg_DirEle"
            Tab(1).Control(13)=   "pnl_Cyg_Celula"
            Tab(1).Control(14)=   "pnl_Cyg_LugNac"
            Tab(1).Control(15)=   "pnl_Cyg_Profes"
            Tab(1).Control(16)=   "pnl_Cyg_NivEst"
            Tab(1).Control(17)=   "pnl_Cyg_Paises"
            Tab(1).Control(18)=   "pnl_Cyg_FecNac"
            Tab(1).Control(19)=   "tab_DatCyg"
            Tab(1).ControlCount=   20
            TabCaption(2)   =   "Datos Crediticios"
            TabPicture(2)   =   "OpeTra_frm_010.frx":0800
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "pnl_Cre_ComVta"
            Tab(2).Control(1)=   "pnl_Cre_TipMon"
            Tab(2).Control(2)=   "pnl_Cre_ApoPro"
            Tab(2).Control(3)=   "pnl_Cre_MonSol_Dol"
            Tab(2).Control(4)=   "pnl_Cre_MonSol_Sol"
            Tab(2).Control(5)=   "pnl_Cre_MonSol_MPr"
            Tab(2).Control(6)=   "pnl_Cre_CuoFij_Dol"
            Tab(2).Control(7)=   "pnl_Cre_CuoIni_Dol"
            Tab(2).Control(8)=   "pnl_Cre_CuoFin_Dol"
            Tab(2).Control(9)=   "pnl_Cre_PlaApr"
            Tab(2).Control(10)=   "pnl_Cre_CuoExt"
            Tab(2).Control(11)=   "pnl_Cre_PerGra"
            Tab(2).Control(12)=   "pnl_Cre_ILDTit"
            Tab(2).Control(13)=   "pnl_Cre_ILDCyg"
            Tab(2).Control(14)=   "pnl_Cre_TCaDol"
            Tab(2).Control(15)=   "pnl_Cre_TCaMPr"
            Tab(2).Control(16)=   "pnl_Cre_MonApr_Dol"
            Tab(2).Control(17)=   "pnl_Cre_MonApr_Sol"
            Tab(2).Control(18)=   "pnl_Cre_MonApr_MPr"
            Tab(2).Control(19)=   "SSPanel39"
            Tab(2).Control(20)=   "pnl_Cre_CuoRen"
            Tab(2).Control(21)=   "pnl_Cre_TasInt"
            Tab(2).Control(22)=   "pnl_Cre_CuoFij_Sol"
            Tab(2).Control(23)=   "pnl_Cre_CuoIni_Sol"
            Tab(2).Control(24)=   "pnl_Cre_CuoFin_Sol"
            Tab(2).Control(25)=   "pnl_Cre_CuoFij_MPr"
            Tab(2).Control(26)=   "pnl_Cre_CuoIni_MPr"
            Tab(2).Control(27)=   "pnl_Cre_CuoFin_MPr"
            Tab(2).Control(28)=   "lbl_NomGlo(201)"
            Tab(2).Control(29)=   "lbl_NomGlo(200)"
            Tab(2).Control(30)=   "lbl_NomGlo(199)"
            Tab(2).Control(31)=   "lbl_NomGlo(198)"
            Tab(2).Control(32)=   "lbl_NomGlo(197)"
            Tab(2).Control(33)=   "lbl_NomGlo(196)"
            Tab(2).Control(34)=   "lbl_NomGlo(190)"
            Tab(2).Control(35)=   "lbl_NomGlo(131)"
            Tab(2).Control(36)=   "lbl_NomGlo(130)"
            Tab(2).Control(37)=   "lbl_NomGlo(129)"
            Tab(2).Control(38)=   "lbl_NomGlo(128)"
            Tab(2).Control(39)=   "lbl_NomGlo(127)"
            Tab(2).Control(40)=   "lbl_NomGlo(126)"
            Tab(2).Control(41)=   "lbl_NomGlo(125)"
            Tab(2).Control(42)=   "lbl_NomGlo(124)"
            Tab(2).Control(43)=   "lbl_NomGlo(123)"
            Tab(2).Control(44)=   "lbl_NomGlo(122)"
            Tab(2).Control(45)=   "lbl_NomGlo(121)"
            Tab(2).Control(46)=   "lbl_NomGlo(120)"
            Tab(2).Control(47)=   "lbl_NomGlo(119)"
            Tab(2).Control(48)=   "lbl_NomGlo(118)"
            Tab(2).Control(49)=   "lbl_NomGlo(117)"
            Tab(2).Control(50)=   "lbl_NomGlo(116)"
            Tab(2).Control(51)=   "lbl_NomGlo(115)"
            Tab(2).Control(52)=   "lbl_NomGlo(114)"
            Tab(2).Control(53)=   "lbl_NomGlo(113)"
            Tab(2).Control(54)=   "lbl_NomGlo(112)"
            Tab(2).ControlCount=   55
            TabCaption(3)   =   "Datos Tasación"
            TabPicture(3)   =   "OpeTra_frm_010.frx":081C
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "lbl_NomGlo(132)"
            Tab(3).Control(1)=   "lbl_NomGlo(133)"
            Tab(3).Control(2)=   "lbl_NomGlo(134)"
            Tab(3).Control(3)=   "lbl_NomGlo(135)"
            Tab(3).Control(4)=   "lbl_NomGlo(136)"
            Tab(3).Control(5)=   "lbl_NomGlo(137)"
            Tab(3).Control(6)=   "lbl_NomGlo(138)"
            Tab(3).Control(7)=   "lbl_NomGlo(139)"
            Tab(3).Control(8)=   "lbl_NomGlo(140)"
            Tab(3).Control(9)=   "lbl_NomGlo(141)"
            Tab(3).Control(10)=   "lbl_NomGlo(142)"
            Tab(3).Control(11)=   "SSPanel46"
            Tab(3).Control(12)=   "SSPanel45"
            Tab(3).Control(13)=   "pnl_Tas_TotATe"
            Tab(3).Control(14)=   "pnl_Tas_TotACo"
            Tab(3).Control(15)=   "pnl_Tas_TotVRe"
            Tab(3).Control(16)=   "pnl_Tas_TotVCo"
            Tab(3).Control(17)=   "pnl_Tas_ATeDep"
            Tab(3).Control(18)=   "pnl_Tas_ACoDep"
            Tab(3).Control(19)=   "pnl_Tas_VReDep"
            Tab(3).Control(20)=   "pnl_Tas_VCoDep"
            Tab(3).Control(21)=   "pnl_Tas_ATeEs2"
            Tab(3).Control(22)=   "pnl_Tas_ACoEs2"
            Tab(3).Control(23)=   "pnl_Tas_VReEs2"
            Tab(3).Control(24)=   "pnl_Tas_VCoEs2"
            Tab(3).Control(25)=   "pnl_Tas_ATeEs1"
            Tab(3).Control(26)=   "pnl_Tas_ACoEs1"
            Tab(3).Control(27)=   "pnl_Tas_VReEs1"
            Tab(3).Control(28)=   "pnl_Tas_VCoEs1"
            Tab(3).Control(29)=   "SSPanel44"
            Tab(3).Control(30)=   "SSPanel43"
            Tab(3).Control(31)=   "SSPanel42"
            Tab(3).Control(32)=   "SSPanel41"
            Tab(3).Control(33)=   "pnl_Tas_AreTer"
            Tab(3).Control(34)=   "pnl_Tas_AreCon"
            Tab(3).Control(35)=   "pnl_Tas_ValRea"
            Tab(3).Control(36)=   "pnl_Tas_NomPer"
            Tab(3).Control(37)=   "pnl_Tas_ValCom"
            Tab(3).Control(38)=   "pnl_Tas_FecEva"
            Tab(3).Control(39)=   "pnl_Tas_NumInf"
            Tab(3).Control(40)=   "pnl_Tas_EmpPer"
            Tab(3).Control(41)=   "pnl_Tas_FecEmi"
            Tab(3).Control(42)=   "txt_Tas_Observ"
            Tab(3).ControlCount=   43
            TabCaption(4)   =   "Datos Seguro"
            TabPicture(4)   =   "OpeTra_frm_010.frx":0838
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "lbl_NomGlo(143)"
            Tab(4).Control(1)=   "lbl_NomGlo(144)"
            Tab(4).Control(2)=   "lbl_NomGlo(145)"
            Tab(4).Control(3)=   "lbl_NomGlo(146)"
            Tab(4).Control(4)=   "lbl_NomGlo(147)"
            Tab(4).Control(5)=   "lbl_NomGlo(148)"
            Tab(4).Control(6)=   "lbl_NomGlo(149)"
            Tab(4).Control(7)=   "lbl_NomGlo(150)"
            Tab(4).Control(8)=   "lbl_NomGlo(151)"
            Tab(4).Control(9)=   "lbl_NomGlo(152)"
            Tab(4).Control(10)=   "lbl_NomGlo(153)"
            Tab(4).Control(11)=   "lbl_NomGlo(154)"
            Tab(4).Control(12)=   "lbl_NomGlo(155)"
            Tab(4).Control(13)=   "lbl_NomGlo(156)"
            Tab(4).Control(14)=   "lbl_NomGlo(157)"
            Tab(4).Control(15)=   "lbl_NomGlo(158)"
            Tab(4).Control(16)=   "lbl_NomGlo(159)"
            Tab(4).Control(17)=   "SSPanel49"
            Tab(4).Control(18)=   "SSPanel47"
            Tab(4).Control(19)=   "pnl_Seg_FoiViv"
            Tab(4).Control(20)=   "pnl_Seg_AplViv"
            Tab(4).Control(21)=   "pnl_Seg_EvaViv"
            Tab(4).Control(22)=   "pnl_Seg_InfViv"
            Tab(4).Control(23)=   "pnl_Seg_SegViv"
            Tab(4).Control(24)=   "pnl_Seg_EmiViv"
            Tab(4).Control(25)=   "pnl_Seg_PolViv"
            Tab(4).Control(26)=   "pnl_Seg_PolCyg"
            Tab(4).Control(27)=   "pnl_Seg_FoiPre"
            Tab(4).Control(28)=   "pnl_Seg_AplPre"
            Tab(4).Control(29)=   "pnl_Seg_EvaPre"
            Tab(4).Control(30)=   "pnl_Seg_InfPre"
            Tab(4).Control(31)=   "pnl_Seg_SegPre"
            Tab(4).Control(32)=   "pnl_Seg_EmiTit"
            Tab(4).Control(33)=   "pnl_Seg_PolTit"
            Tab(4).Control(34)=   "txt_Seg_ObsPol"
            Tab(4).Control(35)=   "txt_Seg_ObsEva"
            Tab(4).ControlCount=   36
            TabCaption(5)   =   "Datos Legales"
            TabPicture(5)   =   "OpeTra_frm_010.frx":0854
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "txt_Leg_ObsBlq"
            Tab(5).Control(1)=   "txt_Leg_InfLeg"
            Tab(5).Control(2)=   "pnl_Leg_RepLeg"
            Tab(5).Control(3)=   "pnl_Leg_Notari"
            Tab(5).Control(4)=   "pnl_Leg_DocReg"
            Tab(5).Control(5)=   "pnl_Leg_AprCom"
            Tab(5).Control(6)=   "pnl_Leg_FirCon"
            Tab(5).Control(7)=   "pnl_Leg_FecBlq"
            Tab(5).Control(8)=   "lbl_NomGlo(167)"
            Tab(5).Control(9)=   "lbl_NomGlo(166)"
            Tab(5).Control(10)=   "lbl_NomGlo(165)"
            Tab(5).Control(11)=   "lbl_NomGlo(164)"
            Tab(5).Control(12)=   "lbl_NomGlo(163)"
            Tab(5).Control(13)=   "lbl_NomGlo(162)"
            Tab(5).Control(14)=   "lbl_NomGlo(161)"
            Tab(5).Control(15)=   "lbl_NomGlo(160)"
            Tab(5).ControlCount=   16
            TabCaption(6)   =   "Datos COFIDE"
            TabPicture(6)   =   "OpeTra_frm_010.frx":0870
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "pnl_Cof_NumCar"
            Tab(6).Control(1)=   "pnl_Cof_NumOpe"
            Tab(6).Control(2)=   "pnl_Cof_FecEmi"
            Tab(6).Control(3)=   "pnl_Cof_FecVal"
            Tab(6).Control(4)=   "pnl_Cof_Import"
            Tab(6).Control(5)=   "pnl_Cof_TipMon"
            Tab(6).Control(6)=   "pnl_Cof_NomBan"
            Tab(6).Control(7)=   "pnl_Cof_NumCta"
            Tab(6).Control(8)=   "pnl_Cof_TasInt"
            Tab(6).Control(9)=   "pnl_Cof_TasCom"
            Tab(6).Control(10)=   "lbl_NomGlo(192)"
            Tab(6).Control(11)=   "lbl_NomGlo(191)"
            Tab(6).Control(12)=   "lbl_NomGlo(175)"
            Tab(6).Control(13)=   "lbl_NomGlo(174)"
            Tab(6).Control(14)=   "lbl_NomGlo(173)"
            Tab(6).Control(15)=   "lbl_NomGlo(172)"
            Tab(6).Control(16)=   "lbl_NomGlo(171)"
            Tab(6).Control(17)=   "lbl_NomGlo(170)"
            Tab(6).Control(18)=   "lbl_NomGlo(169)"
            Tab(6).Control(19)=   "lbl_NomGlo(168)"
            Tab(6).ControlCount=   20
            TabCaption(7)   =   "Datos Inmueble"
            TabPicture(7)   =   "OpeTra_frm_010.frx":088C
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "lbl_NomGlo(176)"
            Tab(7).Control(1)=   "lbl_NomGlo(177)"
            Tab(7).Control(2)=   "lbl_NomGlo(178)"
            Tab(7).Control(3)=   "lbl_NomGlo(179)"
            Tab(7).Control(4)=   "lbl_NomGlo(180)"
            Tab(7).Control(5)=   "lbl_NomGlo(181)"
            Tab(7).Control(6)=   "lbl_NomGlo(182)"
            Tab(7).Control(7)=   "SSPanel4"
            Tab(7).Control(8)=   "pnl_Inm_JurRep"
            Tab(7).Control(9)=   "pnl_Inm_JurDir"
            Tab(7).Control(10)=   "pnl_Inm_JurEmp"
            Tab(7).Control(11)=   "pnl_Inm_NatCyg"
            Tab(7).Control(12)=   "pnl_Inm_NatTit"
            Tab(7).Control(13)=   "pnl_Inm_TipPro"
            Tab(7).Control(14)=   "pnl_Inm_Direcc"
            Tab(7).ControlCount=   15
            TabCaption(8)   =   "Autoriz. Desemb."
            TabPicture(8)   =   "OpeTra_frm_010.frx":08A8
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "txt_Aut_Observ"
            Tab(8).Control(1)=   "pnl_Aut_FueFin"
            Tab(8).Control(2)=   "pnl_Aut_BonoBP"
            Tab(8).Control(3)=   "pnl_Aut_FecDes"
            Tab(8).Control(4)=   "lbl_NomGlo(195)"
            Tab(8).Control(5)=   "lbl_NomGlo(194)"
            Tab(8).Control(6)=   "lbl_NomGlo(193)"
            Tab(8).Control(7)=   "lbl_NomGlo(183)"
            Tab(8).ControlCount=   8
            Begin VB.TextBox txt_Aut_Observ 
               Height          =   3525
               Left            =   -73200
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   335
               Text            =   "OpeTra_frm_010.frx":08C4
               Top             =   1410
               Width           =   10755
            End
            Begin VB.TextBox txt_Tas_Observ 
               Height          =   1125
               Left            =   -73200
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Top             =   3810
               Width           =   10725
            End
            Begin VB.TextBox txt_Seg_ObsEva 
               Height          =   615
               Left            =   -73200
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   24
               Text            =   "OpeTra_frm_010.frx":08C8
               Top             =   3690
               Width           =   10755
            End
            Begin VB.TextBox txt_Seg_ObsPol 
               Height          =   615
               Left            =   -73200
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Text            =   "OpeTra_frm_010.frx":08CC
               Top             =   4320
               Width           =   10755
            End
            Begin VB.TextBox txt_Leg_ObsBlq 
               Height          =   555
               Left            =   -73200
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Text            =   "OpeTra_frm_010.frx":08D0
               Top             =   4350
               Width           =   10755
            End
            Begin VB.TextBox txt_Leg_InfLeg 
               Height          =   1935
               Left            =   -73200
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   21
               Text            =   "OpeTra_frm_010.frx":08D4
               Top             =   420
               Width           =   10755
            End
            Begin TabDlg.SSTab tab_DatCli 
               Height          =   2025
               Left            =   60
               TabIndex        =   26
               Top             =   2880
               Width           =   12435
               _ExtentX        =   21934
               _ExtentY        =   3572
               _Version        =   393216
               Style           =   1
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               TabCaption(0)   =   "Actividad Principal"
               TabPicture(0)   =   "OpeTra_frm_010.frx":08D8
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "lbl_NomGlo(11)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "lbl_NomGlo(10)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "grd_Tit_ActPri"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "pnl_Tit_OcuPri"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).ControlCount=   4
               TabCaption(1)   =   "Actividad Secundaria"
               TabPicture(1)   =   "OpeTra_frm_010.frx":08F4
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "pnl_Tit_OcuSec"
               Tab(1).Control(1)=   "grd_Tit_ActSec"
               Tab(1).Control(2)=   "lbl_NomGlo(20)"
               Tab(1).Control(3)=   "lbl_NomGlo(21)"
               Tab(1).ControlCount=   4
               Begin Threed.SSPanel pnl_Tit_OcuPri 
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   27
                  Top             =   390
                  Width           =   4095
                  _Version        =   65536
                  _ExtentX        =   7223
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin MSFlexGridLib.MSFlexGrid grd_Tit_ActPri 
                  Height          =   1245
                  Left            =   1680
                  TabIndex        =   28
                  Top             =   720
                  Width           =   10665
                  _ExtentX        =   18812
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
               Begin Threed.SSPanel pnl_Tit_OcuSec 
                  Height          =   315
                  Left            =   -73290
                  TabIndex        =   29
                  Top             =   390
                  Width           =   4215
                  _Version        =   65536
                  _ExtentX        =   7435
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin MSFlexGridLib.MSFlexGrid grd_Tit_ActSec 
                  Height          =   1245
                  Left            =   -73320
                  TabIndex        =   30
                  Top             =   720
                  Width           =   10665
                  _ExtentX        =   18812
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
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Index           =   10
                  Left            =   120
                  TabIndex        =   34
                  Top             =   390
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Datos Actividad:"
                  Height          =   285
                  Index           =   11
                  Left            =   120
                  TabIndex        =   33
                  Top             =   720
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Index           =   21
                  Left            =   -74880
                  TabIndex        =   32
                  Top             =   390
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Datos Actividad:"
                  Height          =   285
                  Index           =   20
                  Left            =   -74880
                  TabIndex        =   31
                  Top             =   720
                  Width           =   1275
               End
            End
            Begin Threed.SSPanel pnl_Cre_ComVta 
               Height          =   315
               Left            =   -73200
               TabIndex        =   35
               Top             =   750
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_TipMon 
               Height          =   315
               Left            =   -73200
               TabIndex        =   36
               Top             =   420
               Width           =   3855
               _Version        =   65536
               _ExtentX        =   6800
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "NUEVOS SOLES"
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
            Begin Threed.SSPanel pnl_Tit_FecNac 
               Height          =   315
               Left            =   1800
               TabIndex        =   37
               Top             =   420
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tit_Paises 
               Height          =   315
               Left            =   1800
               TabIndex        =   38
               Top             =   750
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tit_EstCiv 
               Height          =   315
               Left            =   1800
               TabIndex        =   39
               Top             =   1410
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tit_NivEst 
               Height          =   315
               Left            =   1800
               TabIndex        =   40
               Top             =   1740
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tit_Profes 
               Height          =   315
               Left            =   8070
               TabIndex        =   41
               Top             =   1740
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tit_LugNac 
               Height          =   315
               Left            =   1800
               TabIndex        =   42
               Top             =   1080
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tit_Celula 
               Height          =   315
               Left            =   8070
               TabIndex        =   43
               Top             =   420
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_Telefo 
               Height          =   315
               Left            =   8070
               TabIndex        =   44
               Top             =   750
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_Direcc 
               Height          =   615
               Left            =   1800
               TabIndex        =   45
               Top             =   2070
               Width           =   10725
               _Version        =   65536
               _ExtentX        =   18918
               _ExtentY        =   1085
               _StockProps     =   15
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
               Alignment       =   0
            End
            Begin Threed.SSPanel pnl_Tit_DirEle 
               Height          =   315
               Left            =   8070
               TabIndex        =   46
               Top             =   1080
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tit_RegCyg 
               Height          =   315
               Left            =   8070
               TabIndex        =   47
               Top             =   1410
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel SSPanel30 
               Height          =   90
               Left            =   30
               TabIndex        =   48
               Top             =   2730
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin TabDlg.SSTab tab_DatCyg 
               Height          =   2025
               Left            =   -74940
               TabIndex        =   49
               Top             =   2880
               Width           =   12435
               _ExtentX        =   21934
               _ExtentY        =   3572
               _Version        =   393216
               Style           =   1
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               TabCaption(0)   =   "Actividad Principal"
               TabPicture(0)   =   "OpeTra_frm_010.frx":0910
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "lbl_NomGlo(25)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "lbl_NomGlo(24)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "grd_Cyg_ActPri"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "pnl_Cyg_OcuPri"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).ControlCount=   4
               TabCaption(1)   =   "Actividad Secundaria"
               TabPicture(1)   =   "OpeTra_frm_010.frx":092C
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "pnl_Cyg_OcuSec"
               Tab(1).Control(1)=   "grd_Cyg_ActSec"
               Tab(1).Control(2)=   "lbl_NomGlo(108)"
               Tab(1).Control(3)=   "lbl_NomGlo(107)"
               Tab(1).ControlCount=   4
               Begin Threed.SSPanel pnl_Cyg_OcuPri 
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   50
                  Top             =   390
                  Width           =   4095
                  _Version        =   65536
                  _ExtentX        =   7223
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin MSFlexGridLib.MSFlexGrid grd_Cyg_ActPri 
                  Height          =   1245
                  Left            =   1680
                  TabIndex        =   51
                  Top             =   720
                  Width           =   10665
                  _ExtentX        =   18812
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
               Begin Threed.SSPanel pnl_Ocupac 
                  Height          =   315
                  Index           =   3
                  Left            =   -73290
                  TabIndex        =   52
                  Top             =   390
                  Width           =   4215
                  _Version        =   65536
                  _ExtentX        =   7435
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   1245
                  Index           =   3
                  Left            =   -73320
                  TabIndex        =   53
                  Top             =   720
                  Width           =   10665
                  _ExtentX        =   18812
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
               Begin Threed.SSPanel pnl_Cyg_OcuSec 
                  Height          =   315
                  Left            =   -73290
                  TabIndex        =   54
                  Top             =   390
                  Width           =   4095
                  _Version        =   65536
                  _ExtentX        =   7223
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin MSFlexGridLib.MSFlexGrid grd_Cyg_ActSec 
                  Height          =   1245
                  Left            =   -73320
                  TabIndex        =   55
                  Top             =   720
                  Width           =   10665
                  _ExtentX        =   18812
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
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Datos Actividad:"
                  Height          =   285
                  Index           =   22
                  Left            =   -74880
                  TabIndex        =   61
                  Top             =   720
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Index           =   23
                  Left            =   -74880
                  TabIndex        =   60
                  Top             =   390
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Datos Actividad:"
                  Height          =   285
                  Index           =   24
                  Left            =   120
                  TabIndex        =   59
                  Top             =   720
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Index           =   25
                  Left            =   120
                  TabIndex        =   58
                  Top             =   390
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Datos Actividad:"
                  Height          =   285
                  Index           =   107
                  Left            =   -74880
                  TabIndex        =   57
                  Top             =   720
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Index           =   108
                  Left            =   -74880
                  TabIndex        =   56
                  Top             =   390
                  Width           =   1275
               End
            End
            Begin Threed.SSPanel pnl_Cyg_FecNac 
               Height          =   315
               Left            =   -73200
               TabIndex        =   62
               Top             =   1080
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Cyg_Paises 
               Height          =   315
               Left            =   -73200
               TabIndex        =   63
               Top             =   1410
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Cyg_NivEst 
               Height          =   315
               Left            =   -73200
               TabIndex        =   64
               Top             =   2070
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Cyg_Profes 
               Height          =   315
               Left            =   -66930
               TabIndex        =   65
               Top             =   2070
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Cyg_LugNac 
               Height          =   315
               Left            =   -73200
               TabIndex        =   66
               Top             =   1740
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Cyg_Celula 
               Height          =   315
               Left            =   -66930
               TabIndex        =   67
               Top             =   1080
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cyg_DirEle 
               Height          =   315
               Left            =   -66930
               TabIndex        =   68
               Top             =   1740
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel SSPanel18 
               Height          =   90
               Left            =   -74970
               TabIndex        =   69
               Top             =   2730
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin Threed.SSPanel pnl_Cyg_DocIde 
               Height          =   315
               Left            =   -73200
               TabIndex        =   70
               Top             =   420
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Cyg_ApeNom 
               Height          =   315
               Left            =   -73200
               TabIndex        =   71
               Top             =   750
               Width           =   10725
               _Version        =   65536
               _ExtentX        =   18918
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Cre_ApoPro 
               Height          =   315
               Left            =   -73200
               TabIndex        =   72
               Top             =   1080
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_MonSol_Dol 
               Height          =   315
               Left            =   -73200
               TabIndex        =   73
               Top             =   1410
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_MonSol_Sol 
               Height          =   315
               Left            =   -68970
               TabIndex        =   74
               Top             =   1410
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_MonSol_MPr 
               Height          =   315
               Left            =   -64830
               TabIndex        =   75
               Top             =   1440
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_CuoFij_Dol 
               Height          =   315
               Left            =   -73200
               TabIndex        =   76
               Top             =   2220
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_CuoIni_Dol 
               Height          =   315
               Left            =   -73200
               TabIndex        =   77
               Top             =   2550
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_CuoFin_Dol 
               Height          =   315
               Left            =   -73200
               TabIndex        =   78
               Top             =   2880
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_PlaApr 
               Height          =   315
               Left            =   -73200
               TabIndex        =   79
               Top             =   3210
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_CuoExt 
               Height          =   315
               Left            =   -73200
               TabIndex        =   80
               Top             =   3540
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_PerGra 
               Height          =   315
               Left            =   -73200
               TabIndex        =   81
               Top             =   3870
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_ILDTit 
               Height          =   315
               Left            =   -73200
               TabIndex        =   82
               Top             =   4200
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_ILDCyg 
               Height          =   315
               Left            =   -68970
               TabIndex        =   83
               Top             =   4200
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_TCaDol 
               Height          =   315
               Left            =   -73200
               TabIndex        =   84
               Top             =   4530
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_TCaMPr 
               Height          =   315
               Left            =   -68970
               TabIndex        =   85
               Top             =   4530
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_MonApr_Dol 
               Height          =   315
               Left            =   -73200
               TabIndex        =   86
               Top             =   1890
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_MonApr_Sol 
               Height          =   315
               Left            =   -68970
               TabIndex        =   87
               Top             =   1890
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_MonApr_MPr 
               Height          =   315
               Left            =   -64830
               TabIndex        =   88
               Top             =   1890
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel SSPanel39 
               Height          =   90
               Left            =   -74970
               TabIndex        =   89
               Top             =   1770
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin Threed.SSPanel pnl_Cre_CuoRen 
               Height          =   315
               Left            =   -64830
               TabIndex        =   90
               Top             =   4200
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Tas_FecEmi 
               Height          =   315
               Left            =   -66120
               TabIndex        =   91
               Top             =   420
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_EmpPer 
               Height          =   315
               Left            =   -73200
               TabIndex        =   92
               Top             =   420
               Width           =   3315
               _Version        =   65536
               _ExtentX        =   5847
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_NumInf 
               Height          =   315
               Left            =   -73200
               TabIndex        =   93
               Top             =   750
               Width           =   3315
               _Version        =   65536
               _ExtentX        =   5847
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_FecEva 
               Height          =   315
               Left            =   -66120
               TabIndex        =   94
               Top             =   750
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_ValCom 
               Height          =   315
               Left            =   -73200
               TabIndex        =   95
               Top             =   1890
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_NomPer 
               Height          =   315
               Left            =   -73200
               TabIndex        =   96
               Top             =   1080
               Width           =   3315
               _Version        =   65536
               _ExtentX        =   5847
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_ValRea 
               Height          =   315
               Left            =   -71520
               TabIndex        =   97
               Top             =   1890
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_AreCon 
               Height          =   315
               Left            =   -68160
               TabIndex        =   98
               Top             =   1890
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_AreTer 
               Height          =   315
               Left            =   -69840
               TabIndex        =   99
               Top             =   1890
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel SSPanel41 
               Height          =   285
               Left            =   -73200
               TabIndex        =   100
               Top             =   1560
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Valor Comerc. US$"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel42 
               Height          =   285
               Left            =   -71520
               TabIndex        =   101
               Top             =   1560
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Valor Fabricac. US$"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel43 
               Height          =   285
               Left            =   -69840
               TabIndex        =   102
               Top             =   1560
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Area Terreno m2"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel SSPanel44 
               Height          =   285
               Left            =   -68160
               TabIndex        =   103
               Top             =   1560
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Area Constr. m2"
               ForeColor       =   16777215
               BackColor       =   32768
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
            Begin Threed.SSPanel pnl_Tas_VCoEs1 
               Height          =   315
               Left            =   -73200
               TabIndex        =   104
               Top             =   2220
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_VReEs1 
               Height          =   315
               Left            =   -71520
               TabIndex        =   105
               Top             =   2220
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_ACoEs1 
               Height          =   315
               Left            =   -68160
               TabIndex        =   106
               Top             =   2220
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_ATeEs1 
               Height          =   315
               Left            =   -69840
               TabIndex        =   107
               Top             =   2220
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_VCoEs2 
               Height          =   315
               Left            =   -73200
               TabIndex        =   108
               Top             =   2550
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_VReEs2 
               Height          =   315
               Left            =   -71520
               TabIndex        =   109
               Top             =   2550
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_ACoEs2 
               Height          =   315
               Left            =   -68160
               TabIndex        =   110
               Top             =   2550
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_ATeEs2 
               Height          =   315
               Left            =   -69840
               TabIndex        =   111
               Top             =   2550
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_VCoDep 
               Height          =   315
               Left            =   -73200
               TabIndex        =   112
               Top             =   2880
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_VReDep 
               Height          =   315
               Left            =   -71520
               TabIndex        =   113
               Top             =   2880
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_ACoDep 
               Height          =   315
               Left            =   -68160
               TabIndex        =   114
               Top             =   2880
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_ATeDep 
               Height          =   315
               Left            =   -69840
               TabIndex        =   115
               Top             =   2880
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_TotVCo 
               Height          =   315
               Left            =   -73200
               TabIndex        =   116
               Top             =   3330
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_TotVRe 
               Height          =   315
               Left            =   -71520
               TabIndex        =   117
               Top             =   3330
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_TotACo 
               Height          =   315
               Left            =   -68160
               TabIndex        =   118
               Top             =   3330
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Tas_TotATe 
               Height          =   315
               Left            =   -69840
               TabIndex        =   119
               Top             =   3330
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel SSPanel45 
               Height          =   90
               Left            =   -74970
               TabIndex        =   120
               Top             =   1440
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin Threed.SSPanel SSPanel46 
               Height          =   90
               Left            =   -74970
               TabIndex        =   121
               Top             =   3690
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin Threed.SSPanel pnl_Seg_PolTit 
               Height          =   315
               Left            =   -73200
               TabIndex        =   122
               Top             =   1410
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Seg_EmiTit 
               Height          =   315
               Left            =   -66150
               TabIndex        =   123
               Top             =   1410
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "01/10/2004"
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
            End
            Begin Threed.SSPanel pnl_Seg_SegPre 
               Height          =   315
               Left            =   -73200
               TabIndex        =   124
               Top             =   420
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Seg_InfPre 
               Height          =   315
               Left            =   -73200
               TabIndex        =   125
               Top             =   750
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Seg_EvaPre 
               Height          =   315
               Left            =   -66150
               TabIndex        =   126
               Top             =   750
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "01/10/2004"
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
            End
            Begin Threed.SSPanel pnl_Seg_AplPre 
               Height          =   315
               Left            =   -73200
               TabIndex        =   127
               Top             =   1080
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.9812345 "
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
            Begin Threed.SSPanel pnl_Seg_FoiPre 
               Height          =   315
               Left            =   -66150
               TabIndex        =   128
               Top             =   1080
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.9812345 "
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
            Begin Threed.SSPanel pnl_Seg_PolCyg 
               Height          =   315
               Left            =   -73200
               TabIndex        =   129
               Top             =   1740
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Seg_PolViv 
               Height          =   315
               Left            =   -73200
               TabIndex        =   130
               Top             =   3210
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Seg_EmiViv 
               Height          =   315
               Left            =   -66150
               TabIndex        =   131
               Top             =   3210
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "01/10/2004"
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
            End
            Begin Threed.SSPanel pnl_Seg_SegViv 
               Height          =   315
               Left            =   -73200
               TabIndex        =   132
               Top             =   2220
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Seg_InfViv 
               Height          =   315
               Left            =   -73200
               TabIndex        =   133
               Top             =   2550
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Seg_EvaViv 
               Height          =   315
               Left            =   -66150
               TabIndex        =   134
               Top             =   2550
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "01/10/2004"
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
            End
            Begin Threed.SSPanel pnl_Seg_AplViv 
               Height          =   315
               Left            =   -73200
               TabIndex        =   135
               Top             =   2880
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.9812345 "
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
            Begin Threed.SSPanel pnl_Seg_FoiViv 
               Height          =   315
               Left            =   -66150
               TabIndex        =   136
               Top             =   2880
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.9812345 "
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
            Begin Threed.SSPanel SSPanel47 
               Height          =   90
               Left            =   -74970
               TabIndex        =   137
               Top             =   2100
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin Threed.SSPanel SSPanel49 
               Height          =   90
               Left            =   -74970
               TabIndex        =   138
               Top             =   3570
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin Threed.SSPanel pnl_Leg_RepLeg 
               Height          =   315
               Left            =   -73200
               TabIndex        =   139
               Top             =   3360
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
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
            Begin Threed.SSPanel pnl_Leg_Notari 
               Height          =   315
               Left            =   -73200
               TabIndex        =   140
               Top             =   3030
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
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
            Begin Threed.SSPanel pnl_Leg_DocReg 
               Height          =   315
               Left            =   -73200
               TabIndex        =   141
               Top             =   4020
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
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
            Begin Threed.SSPanel pnl_Leg_AprCom 
               Height          =   315
               Left            =   -73200
               TabIndex        =   142
               Top             =   2370
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
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
            Begin Threed.SSPanel pnl_Leg_FirCon 
               Height          =   315
               Left            =   -73200
               TabIndex        =   143
               Top             =   2700
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
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
            Begin Threed.SSPanel pnl_Leg_FecBlq 
               Height          =   315
               Left            =   -73200
               TabIndex        =   144
               Top             =   3690
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
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
            Begin Threed.SSPanel pnl_Cof_NumCar 
               Height          =   315
               Left            =   -73200
               TabIndex        =   145
               Top             =   420
               Width           =   2445
               _Version        =   65536
               _ExtentX        =   4313
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Cof_NumOpe 
               Height          =   315
               Left            =   -73200
               TabIndex        =   146
               Top             =   2070
               Width           =   2445
               _Version        =   65536
               _ExtentX        =   4313
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Cof_FecEmi 
               Height          =   315
               Left            =   -73200
               TabIndex        =   147
               Top             =   750
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
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
            Begin Threed.SSPanel pnl_Cof_FecVal 
               Height          =   315
               Left            =   -73200
               TabIndex        =   148
               Top             =   1080
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
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
            Begin Threed.SSPanel pnl_Cof_Import 
               Height          =   315
               Left            =   -73200
               TabIndex        =   149
               Top             =   2730
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cof_TipMon 
               Height          =   315
               Left            =   -73200
               TabIndex        =   150
               Top             =   2400
               Width           =   3855
               _Version        =   65536
               _ExtentX        =   6800
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "NUEVOS SOLES"
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
            Begin Threed.SSPanel pnl_Cof_NomBan 
               Height          =   315
               Left            =   -73200
               TabIndex        =   151
               Top             =   1410
               Width           =   3315
               _Version        =   65536
               _ExtentX        =   5847
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
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
            Begin Threed.SSPanel pnl_Cof_NumCta 
               Height          =   315
               Left            =   -73200
               TabIndex        =   152
               Top             =   1740
               Width           =   3315
               _Version        =   65536
               _ExtentX        =   5847
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
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
            Begin Threed.SSPanel pnl_Inm_Direcc 
               Height          =   615
               Left            =   -73200
               TabIndex        =   153
               Top             =   420
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   1085
               _StockProps     =   15
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
               Alignment       =   0
            End
            Begin Threed.SSPanel pnl_Inm_TipPro 
               Height          =   315
               Left            =   -73200
               TabIndex        =   154
               Top             =   1050
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Inm_NatTit 
               Height          =   315
               Left            =   -73200
               TabIndex        =   155
               Top             =   1530
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Inm_NatCyg 
               Height          =   315
               Left            =   -73200
               TabIndex        =   156
               Top             =   1860
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Inm_JurEmp 
               Height          =   315
               Left            =   -73200
               TabIndex        =   157
               Top             =   2340
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Inm_JurDir 
               Height          =   615
               Left            =   -73200
               TabIndex        =   158
               Top             =   2670
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   1085
               _StockProps     =   15
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
               Alignment       =   0
            End
            Begin Threed.SSPanel pnl_Inm_JurRep 
               Height          =   315
               Left            =   -73200
               TabIndex        =   159
               Top             =   3300
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel SSPanel4 
               Height          =   90
               Left            =   -74970
               TabIndex        =   160
               Top             =   1410
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin Threed.SSPanel SSPanel3 
               Height          =   90
               Left            =   -74970
               TabIndex        =   161
               Top             =   2220
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
               _ExtentY        =   159
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
            End
            Begin Threed.SSPanel pnl_Cre_TasInt 
               Height          =   315
               Left            =   -64830
               TabIndex        =   406
               Top             =   420
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cof_TasInt 
               Height          =   315
               Left            =   -73200
               TabIndex        =   408
               Top             =   3060
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cof_TasCom 
               Height          =   315
               Left            =   -73200
               TabIndex        =   410
               Top             =   3390
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Aut_FueFin 
               Height          =   315
               Left            =   -73200
               TabIndex        =   412
               Top             =   420
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Aut_BonoBP 
               Height          =   315
               Left            =   -73200
               TabIndex        =   414
               Top             =   750
               Width           =   1485
               _Version        =   65536
               _ExtentX        =   2619
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Aut_FecDes 
               Height          =   315
               Left            =   -73200
               TabIndex        =   416
               Top             =   1080
               Width           =   1485
               _Version        =   65536
               _ExtentX        =   2619
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Cre_CuoFij_Sol 
               Height          =   315
               Left            =   -68970
               TabIndex        =   418
               Top             =   2220
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_CuoIni_Sol 
               Height          =   315
               Left            =   -68970
               TabIndex        =   419
               Top             =   2550
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_CuoFin_Sol 
               Height          =   315
               Left            =   -68970
               TabIndex        =   420
               Top             =   2880
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_CuoFij_MPr 
               Height          =   315
               Left            =   -64830
               TabIndex        =   424
               Top             =   2220
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_CuoIni_MPr 
               Height          =   315
               Left            =   -64830
               TabIndex        =   425
               Top             =   2550
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin Threed.SSPanel pnl_Cre_CuoFin_MPr 
               Height          =   315
               Left            =   -64830
               TabIndex        =   426
               Top             =   2880
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Ultima Cuota MPr.:"
               Height          =   315
               Index           =   201
               Left            =   -66510
               TabIndex        =   429
               Top             =   2880
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Primera Cuota MPr.:"
               Height          =   315
               Index           =   200
               Left            =   -66510
               TabIndex        =   428
               Top             =   2550
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cuota Fija MPr.:"
               Height          =   315
               Index           =   199
               Left            =   -66510
               TabIndex        =   427
               Top             =   2220
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Ultima Cuota S/.:"
               Height          =   315
               Index           =   198
               Left            =   -70650
               TabIndex        =   423
               Top             =   2880
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Primera Cuota S/.:"
               Height          =   315
               Index           =   197
               Left            =   -70650
               TabIndex        =   422
               Top             =   2550
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cuota Fija S/.:"
               Height          =   315
               Index           =   196
               Left            =   -70650
               TabIndex        =   421
               Top             =   2220
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Fecha Desembolso:"
               Height          =   285
               Index           =   195
               Left            =   -74880
               TabIndex        =   417
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Flag Bono Buen Pag.:"
               Height          =   285
               Index           =   194
               Left            =   -74880
               TabIndex        =   415
               Top             =   750
               Width           =   1575
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Fuente Financ.:"
               Height          =   285
               Index           =   193
               Left            =   -74880
               TabIndex        =   413
               Top             =   420
               Width           =   1515
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tasa Comisión:"
               Height          =   315
               Index           =   192
               Left            =   -74880
               TabIndex        =   411
               Top             =   3420
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tasa Interés:"
               Height          =   315
               Index           =   191
               Left            =   -74880
               TabIndex        =   409
               Top             =   3090
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tasa de Interés:"
               Height          =   315
               Index           =   190
               Left            =   -66540
               TabIndex        =   407
               Top             =   420
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Comentarios de Autorización:"
               Height          =   825
               Index           =   183
               Left            =   -74880
               TabIndex        =   336
               Top             =   1410
               Width           =   1305
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Rep. Legal:"
               Height          =   315
               Index           =   182
               Left            =   -74880
               TabIndex        =   334
               Top             =   3300
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Dirección Empresa:"
               Height          =   285
               Index           =   181
               Left            =   -74880
               TabIndex        =   333
               Top             =   2670
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Empresa:"
               Height          =   315
               Index           =   180
               Left            =   -74880
               TabIndex        =   332
               Top             =   2340
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cónyuge:"
               Height          =   315
               Index           =   179
               Left            =   -74880
               TabIndex        =   331
               Top             =   1860
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Titular:"
               Height          =   315
               Index           =   178
               Left            =   -74880
               TabIndex        =   330
               Top             =   1530
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo Propietario:"
               Height          =   315
               Index           =   177
               Left            =   -74880
               TabIndex        =   329
               Top             =   1050
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Dirección:"
               Height          =   285
               Index           =   176
               Left            =   -74880
               TabIndex        =   328
               Top             =   420
               Width           =   1305
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Moneda:"
               Height          =   315
               Index           =   175
               Left            =   -74880
               TabIndex        =   327
               Top             =   2400
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Importe:"
               Height          =   315
               Index           =   174
               Left            =   -74880
               TabIndex        =   326
               Top             =   2760
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión Valor:"
               Height          =   315
               Index           =   173
               Left            =   -74880
               TabIndex        =   325
               Top             =   1080
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión Carta:"
               Height          =   315
               Index           =   172
               Left            =   -74880
               TabIndex        =   324
               Top             =   750
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Operación:"
               Height          =   285
               Index           =   171
               Left            =   -74880
               TabIndex        =   323
               Top             =   2070
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Carta:"
               Height          =   285
               Index           =   170
               Left            =   -74880
               TabIndex        =   322
               Top             =   420
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Banco Transferencia:"
               Height          =   315
               Index           =   169
               Left            =   -74880
               TabIndex        =   321
               Top             =   1410
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número de Cuenta:"
               Height          =   315
               Index           =   168
               Left            =   -74880
               TabIndex        =   320
               Top             =   1740
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Inscrito en:"
               Height          =   315
               Index           =   167
               Left            =   -74880
               TabIndex        =   319
               Top             =   4020
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Firma Minuta:"
               Height          =   315
               Index           =   166
               Left            =   -74880
               TabIndex        =   318
               Top             =   2700
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Aprob. Comité:"
               Height          =   315
               Index           =   165
               Left            =   -74880
               TabIndex        =   317
               Top             =   2370
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Comentarios Bloq.:"
               Height          =   465
               Index           =   164
               Left            =   -74880
               TabIndex        =   316
               Top             =   4350
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Bloqueo Regist.:"
               Height          =   315
               Index           =   163
               Left            =   -74880
               TabIndex        =   315
               Top             =   3690
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Notaria:"
               Height          =   315
               Index           =   162
               Left            =   -74880
               TabIndex        =   314
               Top             =   3030
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Repres. Legal (es):"
               Height          =   315
               Index           =   161
               Left            =   -74880
               TabIndex        =   313
               Top             =   3360
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Informe Legal:"
               Height          =   315
               Index           =   160
               Left            =   -74880
               TabIndex        =   312
               Top             =   420
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo de Aplicación:"
               Height          =   285
               Index           =   159
               Left            =   -74880
               TabIndex        =   311
               Top             =   2880
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Informe:"
               Height          =   285
               Index           =   158
               Left            =   -67470
               TabIndex        =   310
               Top             =   2550
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Informe:"
               Height          =   285
               Index           =   157
               Left            =   -74880
               TabIndex        =   309
               Top             =   2550
               Width           =   1395
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Factor/Importe:"
               Height          =   285
               Index           =   156
               Left            =   -67470
               TabIndex        =   308
               Top             =   2880
               Width           =   1155
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Seguro de Vivienda:"
               Height          =   285
               Index           =   155
               Left            =   -74880
               TabIndex        =   307
               Top             =   2220
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nro Póliza:"
               Height          =   285
               Index           =   154
               Left            =   -74880
               TabIndex        =   306
               Top             =   3210
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión:"
               Height          =   285
               Index           =   153
               Left            =   -67470
               TabIndex        =   305
               Top             =   3210
               Width           =   1095
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nro Póliza (Cyg.):"
               Height          =   285
               Index           =   152
               Left            =   -74880
               TabIndex        =   304
               Top             =   1740
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo de Aplicación:"
               Height          =   285
               Index           =   151
               Left            =   -74880
               TabIndex        =   303
               Top             =   1080
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Informe:"
               Height          =   285
               Index           =   150
               Left            =   -67470
               TabIndex        =   302
               Top             =   750
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Informe:"
               Height          =   285
               Index           =   149
               Left            =   -74880
               TabIndex        =   301
               Top             =   750
               Width           =   1395
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Factor/Importe:"
               Height          =   285
               Index           =   148
               Left            =   -67470
               TabIndex        =   300
               Top             =   1080
               Width           =   1155
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Seguro de Préstamo:"
               Height          =   285
               Index           =   147
               Left            =   -74880
               TabIndex        =   299
               Top             =   420
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión:"
               Height          =   285
               Index           =   146
               Left            =   -67470
               TabIndex        =   298
               Top             =   1410
               Width           =   1095
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nro Póliza (Tit.):"
               Height          =   285
               Index           =   145
               Left            =   -74880
               TabIndex        =   297
               Top             =   1410
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Observaciones de Evaluación:"
               Height          =   435
               Index           =   144
               Left            =   -74880
               TabIndex        =   296
               Top             =   3690
               Width           =   1605
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Observaciones de Tramitación de Póliza:"
               Height          =   435
               Index           =   143
               Left            =   -74880
               TabIndex        =   295
               Top             =   4320
               Width           =   1605
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Totales:"
               Height          =   285
               Index           =   142
               Left            =   -74880
               TabIndex        =   294
               Top             =   3330
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Depósito:"
               Height          =   285
               Index           =   141
               Left            =   -74880
               TabIndex        =   293
               Top             =   2880
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Estacionam. 2:"
               Height          =   285
               Index           =   140
               Left            =   -74880
               TabIndex        =   292
               Top             =   2550
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nombre Perito:"
               Height          =   285
               Index           =   139
               Left            =   -74880
               TabIndex        =   291
               Top             =   1080
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Empresa Peritaje:"
               Height          =   285
               Index           =   138
               Left            =   -74880
               TabIndex        =   290
               Top             =   420
               Width           =   1725
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Inmueble:"
               Height          =   285
               Index           =   137
               Left            =   -74880
               TabIndex        =   289
               Top             =   1890
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Evaluación:"
               Height          =   285
               Index           =   136
               Left            =   -67470
               TabIndex        =   288
               Top             =   750
               Width           =   1155
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Informe:"
               Height          =   285
               Index           =   135
               Left            =   -74880
               TabIndex        =   287
               Top             =   750
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Estacionam. 1:"
               Height          =   285
               Index           =   134
               Left            =   -74880
               TabIndex        =   286
               Top             =   2220
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Observaciones:"
               Height          =   315
               Index           =   133
               Left            =   -74880
               TabIndex        =   285
               Top             =   3810
               Width           =   1365
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión OT:"
               Height          =   315
               Index           =   132
               Left            =   -67470
               TabIndex        =   284
               Top             =   420
               Width           =   1275
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Solic. S/.:"
               Height          =   315
               Index           =   131
               Left            =   -70650
               TabIndex        =   283
               Top             =   1410
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Solic. MPr.:"
               Height          =   315
               Index           =   130
               Left            =   -66510
               TabIndex        =   282
               Top             =   1410
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "ILD Cónyuge S/.:"
               Height          =   315
               Index           =   129
               Left            =   -70650
               TabIndex        =   281
               Top             =   4200
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo Cambio MPr.:"
               Height          =   315
               Index           =   128
               Left            =   -70650
               TabIndex        =   280
               Top             =   4530
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. S/.:"
               Height          =   315
               Index           =   127
               Left            =   -70650
               TabIndex        =   279
               Top             =   1890
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. MPr.:"
               Height          =   315
               Index           =   126
               Left            =   -66510
               TabIndex        =   278
               Top             =   1890
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Relac. Cuota/Renta:"
               Height          =   315
               Index           =   125
               Left            =   -66510
               TabIndex        =   277
               Top             =   4200
               Width           =   1665
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Moneda Préstamo:"
               Height          =   315
               Index           =   124
               Left            =   -74880
               TabIndex        =   276
               Top             =   420
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "V. Compra-Venta US$:"
               Height          =   315
               Index           =   123
               Left            =   -74880
               TabIndex        =   275
               Top             =   750
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Aporte Propio US$:"
               Height          =   315
               Index           =   122
               Left            =   -74880
               TabIndex        =   274
               Top             =   1080
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Solic. US$:"
               Height          =   315
               Index           =   121
               Left            =   -74880
               TabIndex        =   273
               Top             =   1410
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cuota Fija US$:"
               Height          =   315
               Index           =   120
               Left            =   -74880
               TabIndex        =   272
               Top             =   2220
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Primera Cuota US$:"
               Height          =   315
               Index           =   119
               Left            =   -74880
               TabIndex        =   271
               Top             =   2550
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Ultima Cuota US$:"
               Height          =   315
               Index           =   118
               Left            =   -74880
               TabIndex        =   270
               Top             =   2880
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Plazo Aprobado:"
               Height          =   315
               Index           =   117
               Left            =   -74880
               TabIndex        =   269
               Top             =   3210
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cuotas Extraord.:"
               Height          =   315
               Index           =   116
               Left            =   -74880
               TabIndex        =   268
               Top             =   3540
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Período de Gracia:"
               Height          =   315
               Index           =   115
               Left            =   -74880
               TabIndex        =   267
               Top             =   3870
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "ILD Titular S/.:"
               Height          =   315
               Index           =   114
               Left            =   -74880
               TabIndex        =   266
               Top             =   4200
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo Cambio US$:"
               Height          =   315
               Index           =   113
               Left            =   -74880
               TabIndex        =   265
               Top             =   4530
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. US$:"
               Height          =   315
               Index           =   112
               Left            =   -74880
               TabIndex        =   264
               Top             =   1890
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Profesión:"
               Height          =   285
               Index           =   111
               Left            =   6600
               TabIndex        =   263
               Top             =   1740
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Telf. Celular:"
               Height          =   285
               Index           =   110
               Left            =   6600
               TabIndex        =   262
               Top             =   420
               Width           =   1275
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Telf. Casa:"
               Height          =   285
               Index           =   109
               Left            =   6600
               TabIndex        =   261
               Top             =   750
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "E-Mail Personal:"
               Height          =   285
               Index           =   100
               Left            =   6600
               TabIndex        =   260
               Top             =   1080
               Width           =   1245
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Régimen Conyugal:"
               Height          =   285
               Index           =   18
               Left            =   6600
               TabIndex        =   259
               Top             =   1410
               Width           =   1395
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Fecha de Nacimiento:"
               Height          =   285
               Index           =   17
               Left            =   120
               TabIndex        =   258
               Top             =   420
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "País de Nacimiento:"
               Height          =   285
               Index           =   16
               Left            =   120
               TabIndex        =   257
               Top             =   750
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Estado Civil:"
               Height          =   285
               Index           =   15
               Left            =   120
               TabIndex        =   256
               Top             =   1410
               Width           =   1365
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nivel de Estudios:"
               Height          =   285
               Index           =   14
               Left            =   120
               TabIndex        =   255
               Top             =   1740
               Width           =   1455
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Lugar de Nacimiento:"
               Height          =   285
               Index           =   13
               Left            =   120
               TabIndex        =   254
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Dirección:"
               Height          =   285
               Index           =   12
               Left            =   120
               TabIndex        =   253
               Top             =   2070
               Width           =   1515
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Moneda Préstamo:"
               Height          =   315
               Index           =   36
               Left            =   -74880
               TabIndex        =   252
               Top             =   420
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "V. Compra-Venta US$:"
               Height          =   315
               Index           =   37
               Left            =   -74880
               TabIndex        =   251
               Top             =   750
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Fecha de Nacimiento:"
               Height          =   285
               Index           =   0
               Left            =   -74880
               TabIndex        =   250
               Top             =   420
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "País de Nacimiento:"
               Height          =   285
               Index           =   2
               Left            =   -74880
               TabIndex        =   249
               Top             =   750
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Estado Civil:"
               Height          =   285
               Index           =   9
               Left            =   -74880
               TabIndex        =   248
               Top             =   1410
               Width           =   1365
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nivel de Estudios:"
               Height          =   285
               Index           =   8
               Left            =   -74880
               TabIndex        =   247
               Top             =   1740
               Width           =   1455
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Profesión:"
               Height          =   285
               Index           =   6
               Left            =   -68400
               TabIndex        =   246
               Top             =   1740
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Lugar de Nacimiento:"
               Height          =   285
               Index           =   19
               Left            =   -74880
               TabIndex        =   245
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Telf. Celular:"
               Height          =   285
               Index           =   1
               Left            =   -68400
               TabIndex        =   244
               Top             =   420
               Width           =   1275
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Telf. Casa:"
               Height          =   285
               Index           =   3
               Left            =   -68400
               TabIndex        =   243
               Top             =   750
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Dirección:"
               Height          =   285
               Index           =   7
               Left            =   -74880
               TabIndex        =   242
               Top             =   2070
               Width           =   1515
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "E-Mail Personal:"
               Height          =   285
               Index           =   4
               Left            =   -68400
               TabIndex        =   241
               Top             =   1080
               Width           =   1245
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Régimen Conyugal:"
               Height          =   285
               Index           =   5
               Left            =   -68400
               TabIndex        =   240
               Top             =   1410
               Width           =   1395
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Fecha de Nacimiento:"
               Height          =   285
               Index           =   26
               Left            =   -74880
               TabIndex        =   239
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "País de Nacimiento:"
               Height          =   285
               Index           =   27
               Left            =   -74880
               TabIndex        =   238
               Top             =   1410
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nivel de Estudios:"
               Height          =   285
               Index           =   29
               Left            =   -74880
               TabIndex        =   237
               Top             =   2070
               Width           =   1455
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Profesión:"
               Height          =   285
               Index           =   30
               Left            =   -68400
               TabIndex        =   236
               Top             =   2070
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Lugar de Nacimiento:"
               Height          =   285
               Index           =   31
               Left            =   -74880
               TabIndex        =   235
               Top             =   1740
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Telf. Celular:"
               Height          =   285
               Index           =   32
               Left            =   -68400
               TabIndex        =   234
               Top             =   1080
               Width           =   1275
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "E-Mail Personal:"
               Height          =   285
               Index           =   35
               Left            =   -68400
               TabIndex        =   233
               Top             =   1740
               Width           =   1245
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Documento Identidad:"
               Height          =   285
               Index           =   28
               Left            =   -74880
               TabIndex        =   232
               Top             =   420
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Apellidos y Nombres:"
               Height          =   285
               Index           =   34
               Left            =   -74880
               TabIndex        =   231
               Top             =   750
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Aporte Propio US$:"
               Height          =   315
               Index           =   38
               Left            =   -74880
               TabIndex        =   230
               Top             =   1080
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Solic. US$:"
               Height          =   315
               Index           =   39
               Left            =   -74880
               TabIndex        =   229
               Top             =   1410
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Solic. S/.:"
               Height          =   315
               Index           =   40
               Left            =   -70650
               TabIndex        =   228
               Top             =   1410
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Solic. MPr.:"
               Height          =   315
               Index           =   41
               Left            =   -66510
               TabIndex        =   227
               Top             =   1410
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cuota Fija:"
               Height          =   315
               Index           =   42
               Left            =   -74880
               TabIndex        =   226
               Top             =   2220
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Primera Cuota:"
               Height          =   315
               Index           =   43
               Left            =   -74880
               TabIndex        =   225
               Top             =   2550
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Ultima Cuota:"
               Height          =   315
               Index           =   44
               Left            =   -74880
               TabIndex        =   224
               Top             =   2880
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Plazo Aprobado:"
               Height          =   315
               Index           =   45
               Left            =   -74880
               TabIndex        =   223
               Top             =   3210
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cuotas Extraord.:"
               Height          =   315
               Index           =   46
               Left            =   -74880
               TabIndex        =   222
               Top             =   3540
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Período de Gracia:"
               Height          =   315
               Index           =   47
               Left            =   -74880
               TabIndex        =   221
               Top             =   3870
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "ILD Titular S/.:"
               Height          =   315
               Index           =   48
               Left            =   -74880
               TabIndex        =   220
               Top             =   4200
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "ILD Cónyuge S/.:"
               Height          =   315
               Index           =   49
               Left            =   -70650
               TabIndex        =   219
               Top             =   4200
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo Cambio US$:"
               Height          =   315
               Index           =   50
               Left            =   -74880
               TabIndex        =   218
               Top             =   4530
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo Cambio MPr.:"
               Height          =   315
               Index           =   51
               Left            =   -70650
               TabIndex        =   217
               Top             =   4530
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. US$:"
               Height          =   315
               Index           =   52
               Left            =   -74880
               TabIndex        =   216
               Top             =   1890
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. S/.:"
               Height          =   315
               Index           =   53
               Left            =   -70650
               TabIndex        =   215
               Top             =   1890
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. MPr.:"
               Height          =   315
               Index           =   54
               Left            =   -66510
               TabIndex        =   214
               Top             =   1890
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Relac. Cuota/Renta:"
               Height          =   315
               Index           =   55
               Left            =   -66510
               TabIndex        =   213
               Top             =   4200
               Width           =   1665
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Totales:"
               Height          =   285
               Index           =   65
               Left            =   -74880
               TabIndex        =   212
               Top             =   3330
               Width           =   1485
            End
            Begin VB.Line Line1 
               BorderWidth     =   2
               X1              =   -73230
               X2              =   -66510
               Y1              =   3270
               Y2              =   3270
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Depósito:"
               Height          =   285
               Index           =   64
               Left            =   -74880
               TabIndex        =   211
               Top             =   2880
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Estacionam. 2:"
               Height          =   285
               Index           =   63
               Left            =   -74880
               TabIndex        =   210
               Top             =   2550
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nombre Perito:"
               Height          =   285
               Index           =   60
               Left            =   -74880
               TabIndex        =   209
               Top             =   1080
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Empresa Peritaje:"
               Height          =   285
               Index           =   56
               Left            =   -74880
               TabIndex        =   208
               Top             =   420
               Width           =   1725
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Inmueble:"
               Height          =   285
               Index           =   61
               Left            =   -74880
               TabIndex        =   207
               Top             =   1890
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Evaluación:"
               Height          =   285
               Index           =   58
               Left            =   -67470
               TabIndex        =   206
               Top             =   750
               Width           =   1155
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Informe:"
               Height          =   285
               Index           =   59
               Left            =   -74880
               TabIndex        =   205
               Top             =   750
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Estacionam. 1:"
               Height          =   285
               Index           =   62
               Left            =   -74880
               TabIndex        =   204
               Top             =   2220
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Observaciones:"
               Height          =   315
               Index           =   66
               Left            =   -74880
               TabIndex        =   203
               Top             =   3810
               Width           =   1365
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión OT:"
               Height          =   315
               Index           =   57
               Left            =   -67470
               TabIndex        =   202
               Top             =   420
               Width           =   1275
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo de Aplicación:"
               Height          =   285
               Index           =   77
               Left            =   -74880
               TabIndex        =   201
               Top             =   2880
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Informe:"
               Height          =   285
               Index           =   81
               Left            =   -67470
               TabIndex        =   200
               Top             =   2550
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Informe:"
               Height          =   285
               Index           =   76
               Left            =   -74880
               TabIndex        =   199
               Top             =   2550
               Width           =   1395
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Factor/Importe:"
               Height          =   285
               Index           =   80
               Left            =   -67470
               TabIndex        =   198
               Top             =   2880
               Width           =   1155
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Seguro de Vivienda:"
               Height          =   285
               Index           =   75
               Left            =   -74880
               TabIndex        =   197
               Top             =   2220
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nro Póliza:"
               Height          =   285
               Index           =   78
               Left            =   -74880
               TabIndex        =   196
               Top             =   3210
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión:"
               Height          =   285
               Index           =   79
               Left            =   -67470
               TabIndex        =   195
               Top             =   3210
               Width           =   1095
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nro Póliza (Cyg.):"
               Height          =   285
               Index           =   71
               Left            =   -74880
               TabIndex        =   194
               Top             =   1740
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo de Aplicación:"
               Height          =   285
               Index           =   69
               Left            =   -74880
               TabIndex        =   193
               Top             =   1080
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Informe:"
               Height          =   285
               Index           =   74
               Left            =   -67470
               TabIndex        =   192
               Top             =   750
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Informe:"
               Height          =   285
               Index           =   68
               Left            =   -74880
               TabIndex        =   191
               Top             =   750
               Width           =   1395
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Factor/Importe:"
               Height          =   285
               Index           =   73
               Left            =   -67470
               TabIndex        =   190
               Top             =   1080
               Width           =   1155
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Seguro de Préstamo:"
               Height          =   285
               Index           =   67
               Left            =   -74880
               TabIndex        =   189
               Top             =   420
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión:"
               Height          =   285
               Index           =   72
               Left            =   -67470
               TabIndex        =   188
               Top             =   1410
               Width           =   1095
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nro Póliza (Tit.):"
               Height          =   285
               Index           =   70
               Left            =   -74880
               TabIndex        =   187
               Top             =   1410
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Observaciones de Evaluación:"
               Height          =   435
               Index           =   82
               Left            =   -74880
               TabIndex        =   186
               Top             =   3690
               Width           =   1605
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Observaciones de Tramitación de Póliza:"
               Height          =   435
               Index           =   83
               Left            =   -74880
               TabIndex        =   185
               Top             =   4320
               Width           =   1605
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Inscrito en:"
               Height          =   315
               Index           =   90
               Left            =   -74880
               TabIndex        =   184
               Top             =   4020
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Firma Minuta:"
               Height          =   315
               Index           =   86
               Left            =   -74880
               TabIndex        =   183
               Top             =   2700
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Aprob. Comité:"
               Height          =   315
               Index           =   85
               Left            =   -74880
               TabIndex        =   182
               Top             =   2370
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Comentarios Bloq.:"
               Height          =   465
               Index           =   91
               Left            =   -74880
               TabIndex        =   181
               Top             =   4350
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Bloqueo Regist.:"
               Height          =   315
               Index           =   89
               Left            =   -74880
               TabIndex        =   180
               Top             =   3690
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Notaria:"
               Height          =   315
               Index           =   87
               Left            =   -74880
               TabIndex        =   179
               Top             =   3030
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Repres. Legal (es):"
               Height          =   315
               Index           =   88
               Left            =   -74880
               TabIndex        =   178
               Top             =   3360
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Informe Legal:"
               Height          =   315
               Index           =   84
               Left            =   -74880
               TabIndex        =   177
               Top             =   420
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Moneda:"
               Height          =   315
               Index           =   98
               Left            =   -74880
               TabIndex        =   176
               Top             =   2400
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Importe:"
               Height          =   315
               Index           =   99
               Left            =   -74880
               TabIndex        =   175
               Top             =   2760
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión Valor:"
               Height          =   315
               Index           =   94
               Left            =   -74880
               TabIndex        =   174
               Top             =   1080
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión Carta:"
               Height          =   315
               Index           =   93
               Left            =   -74880
               TabIndex        =   173
               Top             =   750
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Operación:"
               Height          =   285
               Index           =   97
               Left            =   -74880
               TabIndex        =   172
               Top             =   2070
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Carta:"
               Height          =   285
               Index           =   92
               Left            =   -74880
               TabIndex        =   171
               Top             =   420
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Banco Transferencia:"
               Height          =   315
               Index           =   95
               Left            =   -74880
               TabIndex        =   170
               Top             =   1410
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número de Cuenta:"
               Height          =   315
               Index           =   96
               Left            =   -74880
               TabIndex        =   169
               Top             =   1740
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Rep. Legal:"
               Height          =   315
               Index           =   106
               Left            =   -74880
               TabIndex        =   168
               Top             =   3300
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Dirección Empresa:"
               Height          =   285
               Index           =   105
               Left            =   -74880
               TabIndex        =   167
               Top             =   2670
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Empresa:"
               Height          =   315
               Index           =   104
               Left            =   -74880
               TabIndex        =   166
               Top             =   2340
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cónyuge:"
               Height          =   315
               Index           =   103
               Left            =   -74880
               TabIndex        =   165
               Top             =   1860
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Titular:"
               Height          =   315
               Index           =   102
               Left            =   -74880
               TabIndex        =   164
               Top             =   1530
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo Propietario:"
               Height          =   315
               Index           =   101
               Left            =   -74880
               TabIndex        =   163
               Top             =   1050
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Dirección:"
               Height          =   285
               Index           =   33
               Left            =   -74880
               TabIndex        =   162
               Top             =   420
               Width           =   1305
            End
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1095
         Left            =   60
         TabIndex        =   337
         Top             =   1590
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   1931
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
            Left            =   1620
            TabIndex        =   338
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-001-04-0001"
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   8850
            TabIndex        =   339
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "31/12/2004"
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
         Begin Threed.SSPanel pnl_EjeVta 
            Height          =   315
            Left            =   8850
            TabIndex        =   340
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_Modali 
            Height          =   315
            Left            =   1620
            TabIndex        =   341
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "BIEN TERMINADO"
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1620
            TabIndex        =   342
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO - MIVIVIENDA"
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
         Begin Threed.SSPanel pnl_IniEva 
            Height          =   315
            Left            =   8850
            TabIndex        =   343
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "31/12/2004"
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
         Begin VB.Label lbl_NomGlo 
            Caption         =   "F. Inicio Evaluac.:"
            Height          =   315
            Index           =   189
            Left            =   7410
            TabIndex        =   349
            Top             =   390
            Width           =   1275
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Producto:"
            Height          =   315
            Index           =   188
            Left            =   60
            TabIndex        =   348
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Modalidad:"
            Height          =   315
            Index           =   187
            Left            =   60
            TabIndex        =   347
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Index           =   186
            Left            =   7410
            TabIndex        =   346
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Index           =   185
            Left            =   7410
            TabIndex        =   345
            Top             =   60
            Width           =   1185
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Index           =   184
            Left            =   60
            TabIndex        =   344
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   350
         Top             =   750
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   1402
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10560
            Picture         =   "OpeTra_frm_010.frx":0948
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11280
            Picture         =   "OpeTra_frm_010.frx":0C52
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12000
            Picture         =   "OpeTra_frm_010.frx":0F5C
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   6210
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   2775
         End
         Begin MSMask.MaskEdBox msk_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   3
            Top             =   390
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "###-###-##-####"
            PromptChar      =   " "
         End
         Begin VB.Label Label20 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   355
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label Label30 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   354
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label31 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   353
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label35 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   4830
            TabIndex        =   352
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Solicitud:"
            Height          =   285
            Left            =   90
            TabIndex        =   351
            Top             =   390
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Des_CreHip_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CofNCo()   As modcal_g_est_CuoCof
Dim l_arr_CofCon()   As modcal_g_est_CuoCof
Dim l_arr_CliNCo()   As modcal_g_est_CuoCli
Dim l_arr_CliCon()   As modcal_g_est_CuoCli

Dim l_str_IniEva           As String

Dim l_str_Cli_CodCiu       As String
Dim l_int_Cli_TipVia       As Integer
Dim l_str_Cli_NomVia       As String
Dim l_str_Cli_NumVia       As String
Dim l_str_Cli_IntDpt       As String
Dim l_int_Cli_TipZon       As Integer
Dim l_str_Cli_NomZon       As String
Dim l_str_Cli_Refere       As String
Dim l_str_Cli_UbiGeo       As String
Dim l_str_Cli_Telefo       As String


Dim l_dbl_Cre_MtoPre       As Double
Dim l_dbl_Cre_TasInt       As Double
Dim l_dbl_Cre_TasCof       As Double
Dim l_dbl_Cre_ComCof       As Double
Dim l_int_Cre_NumCuo       As Integer
Dim l_int_Cre_CuoExt       As Integer
Dim l_int_Cre_PerGra       As Integer
Dim l_dbl_Cre_PreMPr       As Double
Dim l_dbl_Cre_PreSol       As Double
Dim l_dbl_Cre_PreDol       As Double
Dim l_dbl_Cre_ComVta       As Double
Dim l_dbl_Cre_ApoPro       As Double
Dim l_dbl_Cre_TCaDol       As Double
Dim l_dbl_Cre_TCaMPr       As Double

Dim l_str_Inm_UbiGeo       As String
Dim l_str_Inm_CodPry       As String
Dim l_int_Inm_PryMCs       As Integer
Dim l_int_Inm_TipVia       As Integer
Dim l_str_Inm_NomVia       As String
Dim l_str_Inm_NumVia       As String
Dim l_str_Inm_IntDpt       As String
Dim l_int_Inm_TipZon       As Integer
Dim l_str_Inm_NomZon       As String
Dim l_str_Inm_Refere       As String
Dim l_str_Inm_Telefo       As String


Dim l_int_Aut_BonoBP       As Integer
Dim l_str_Aut_FecDes       As String
Dim l_str_Aut_FueFin       As String

Dim l_str_Leg_AprCom       As String

Dim l_dbl_Seg_FoiPre       As Double
Dim l_dbl_Seg_FoiViv       As Double
Dim l_int_Seg_AplPre       As Integer
Dim l_int_Seg_TipSeg       As Integer
Dim l_int_Seg_AplViv       As Integer
Dim l_str_Seg_EmpDes       As String
Dim l_str_Seg_EmpViv       As String

Dim l_str_Tas_EmpPer       As String
Dim l_str_Tas_NomPer       As String
Dim l_dbl_Tas_ValCom       As Double
Dim l_dbl_Tas_ValFab       As Double
Dim l_int_Tas_TipMon       As Integer

Dim l_str_Cof_NumOpe       As String
Dim l_int_Cof_TipMon       As Integer
Dim l_dbl_Cof_MtoDes       As Double

Dim l_str_CliNCo_PriVct    As String
Dim l_str_CliNCo_UltVct    As String
Dim l_str_CliNCo_PrxVct    As String
Dim l_dbl_CliNCo_CuoIni    As Double
Dim l_dbl_CliNCo_CuoFin    As Double
Dim l_dbl_CliNCo_CuoFij    As Double

Dim l_str_CofNCo_PriVct    As String
Dim l_str_CofNCo_UltVct    As String
Dim l_str_CofNCo_PrxVct    As String
Dim l_dbl_CofNCo_CuoIni    As Double
Dim l_dbl_CofNCo_CuoFin    As Double
Dim l_dbl_CofNCo_CuoFij    As Double

Dim l_str_CofCon_PriVct    As String
Dim l_str_CofCon_UltVct    As String
Dim l_str_CofCon_PrxVct    As String
Dim l_dbl_CofCon_CuoIni    As Double
Dim l_dbl_CofCon_CuoFin    As Double
Dim l_dbl_CofCon_CuoFij    As Double


Dim l_dbl_PorNCo        As Double
Dim l_dbl_PorCon        As Double
Dim l_dbl_ImpNCo        As Double
Dim l_dbl_ImpCon        As Double
Dim l_dbl_OtrCar        As Double
Dim l_dbl_SegPre        As Double
Dim l_dbl_SegViv        As Double
Dim l_int_ClaCre        As Integer
Dim l_int_IndITF        As Integer
Dim l_str_NumOpe        As String


Private Sub cmd_Buscar_Click()
   If cmb_TipBus.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Búsqueda.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_TipBus)
      Exit Sub
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      If cmb_TipDoc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmb_TipDoc)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumDoc.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
         txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
      End If
      
      moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_str_TipDoc = cmb_TipDoc.Text
      moddat_g_str_NumDoc = txt_NumDoc.Text
   Else
      If Len(Trim(msk_NumSol.Text)) < 12 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      moddat_g_str_NumSol = msk_NumSol.Text
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
      g_str_Parame = g_str_Parame & "SOLMAE_ENVCRE = 1 "
   Else
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
      g_str_Parame = g_str_Parame & "SOLMAE_ENVCRE = 1 "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe Solicitud en Trámite para la Selección de Búsqueda. ", vbExclamation, modgen_g_con_AteCli
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   Call fs_Buscar_DatGen
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   pnl_EjeVta.Caption = moddat_g_str_EjeVta
   pnl_FecIng.Caption = moddat_g_str_FecIng

   'Validación que se encuentre en Instancia
   If moddat_g_int_InsAct <> modatecli_g_con_Desemb Then
      MsgBox "No se encuentra en Instancia de Autorización de Desembolso.", vbInformation, modgen_g_con_AteCli
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Call fs_ActivaItem(False)
   Call fs_Activa(False)
   
   Call fs_Buscar_SegDet
   
   Call fs_Buscar_InfSol
   Call fs_GenCro
   
   Call fs_ActivaItem(True)
   
   Call gs_SetFocus(grd_CliNCo_Listad)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_Cadena     As String
   
   If MsgBox("¿Está seguro de Generar la Operación Crediticia?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Genera_Operac
   Screen.MousePointer = 0
   
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   r_str_Cadena = r_str_Cadena & "NUMERO DE OPERACION : " & Left(l_str_NumOpe, 3) & "-" & Mid(l_str_NumOpe, 4, 2) & "-" & Right(l_str_NumOpe, 5) & Chr(13)
   r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Cadena = r_str_Cadena & Chr(13)

   modgen_g_str_Mail_Asunto = "GENERACION DE OPERACION (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   modgen_g_str_Mail_Mensaj = r_str_Cadena
   
   frm_EnvMai_01.Show 1
   
   MsgBox "El Número de Operación generado es el : " & Left(l_str_NumOpe, 3) & "-" & Mid(l_str_NumOpe, 4, 2) & "-" & Right(l_str_NumOpe, 5), vbInformation, modgen_g_str_NomPlt
   
   Call cmd_Limpia_Click
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_con_AteCli
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmb_TipBus_Click()
   If cmb_TipBus.ListIndex > -1 Then
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         cmb_TipDoc.Enabled = True
         txt_NumDoc.Enabled = True
         msk_NumSol.Enabled = False
         
         msk_NumSol.Mask = ""
         msk_NumSol.Text = ""
         msk_NumSol.Mask = "###-###-##-####"
         
         Call gs_SetFocus(cmb_TipDoc)
      Else
         cmb_TipDoc.Enabled = False
         txt_NumDoc.Enabled = False
         msk_NumSol.Enabled = True
         
         cmb_TipDoc.ListIndex = -1
         txt_NumDoc.Text = ""
         
         Call gs_SetFocus(msk_NumSol)
      End If
   Else
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      
      msk_NumSol.Enabled = False
   
      cmb_TipDoc.ListIndex = -1
      txt_NumDoc.Text = ""
      msk_NumSol.Mask = ""
      msk_NumSol.Text = ""
      msk_NumSol.Mask = "###-###-##-####"
   End If
End Sub

Private Sub cmb_TipBus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipBus_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub fs_Inicia()
   Call modsis_gs_Carga_TipBus(cmb_TipBus)
   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
   
   grd_Tit_ActPri.ColWidth(0) = 2130
   grd_Tit_ActPri.ColWidth(1) = 8200
   grd_Tit_ActPri.ColAlignment(0) = flexAlignLeftCenter
   grd_Tit_ActPri.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Tit_ActSec.ColWidth(0) = 2130
   grd_Tit_ActSec.ColWidth(1) = 8200
   grd_Tit_ActSec.ColAlignment(0) = flexAlignLeftCenter
   grd_Tit_ActSec.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Cyg_ActPri.ColWidth(0) = 2130
   grd_Cyg_ActPri.ColWidth(1) = 8200
   grd_Cyg_ActPri.ColAlignment(0) = flexAlignLeftCenter
   grd_Cyg_ActPri.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Cyg_ActSec.ColWidth(0) = 2130
   grd_Cyg_ActSec.ColWidth(1) = 8200
   grd_Cyg_ActSec.ColAlignment(0) = flexAlignLeftCenter
   grd_Cyg_ActSec.ColAlignment(1) = flexAlignLeftCenter
   
   'Cliente No Concesional
   grd_CliNCo_Listad.ColWidth(0) = 1110
   grd_CliNCo_Listad.ColWidth(1) = 1380
   grd_CliNCo_Listad.ColWidth(2) = 1395
   grd_CliNCo_Listad.ColWidth(3) = 1390
   grd_CliNCo_Listad.ColWidth(4) = 1385
   grd_CliNCo_Listad.ColWidth(5) = 1380
   grd_CliNCo_Listad.ColWidth(6) = 1370
   grd_CliNCo_Listad.ColWidth(7) = 1375
   grd_CliNCo_Listad.ColWidth(8) = 1380
   
   grd_CliNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(8) = flexAlignRightCenter
   
   'Cliente Concesional
   grd_CliCon_Listad.ColWidth(0) = 1170
   grd_CliCon_Listad.ColWidth(1) = 1560
   grd_CliCon_Listad.ColWidth(2) = 2355
   grd_CliCon_Listad.ColWidth(3) = 2355
   grd_CliCon_Listad.ColWidth(4) = 2355
   grd_CliCon_Listad.ColWidth(5) = 2355
   
   grd_CliCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(5) = flexAlignRightCenter
   
   'Cofide No Concesional
   grd_CofNCo_Listad.ColWidth(0) = 1170
   grd_CofNCo_Listad.ColWidth(1) = 1800
   grd_CofNCo_Listad.ColWidth(2) = 1835
   grd_CofNCo_Listad.ColWidth(3) = 1835
   grd_CofNCo_Listad.ColWidth(4) = 1835
   grd_CofNCo_Listad.ColWidth(5) = 1840
   grd_CofNCo_Listad.ColWidth(6) = 1835
   
   grd_CofNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   
   'Cofide Concesional
   grd_CofCon_Listad.ColWidth(0) = 1170
   grd_CofCon_Listad.ColWidth(1) = 1800
   grd_CofCon_Listad.ColWidth(2) = 1835
   grd_CofCon_Listad.ColWidth(3) = 1835
   grd_CofCon_Listad.ColWidth(4) = 1835
   grd_CofCon_Listad.ColWidth(5) = 1840
   grd_CofCon_Listad.ColWidth(6) = 1835
   
   grd_CofCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CofCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CofCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CofCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CofCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CofCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CofCon_Listad.ColAlignment(6) = flexAlignRightCenter
End Sub

Private Sub msk_NumSol_GotFocus()
   Call gs_SelecTodo(msk_NumSol)
End Sub

Private Sub msk_NumSol_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub txt_Leg_InfLeg_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Leg_ObsBlq_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub fs_Limpia()
   Call fs_ActivaItem(False)
   Call fs_Activa(True)
   
   cmb_TipBus.ListIndex = -1
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   msk_NumSol.Enabled = False

   msk_NumSol.Mask = ""
   msk_NumSol.Text = ""
   msk_NumSol.Mask = "###-###-##-####"
   
   txt_NumDoc.Text = ""
   
   pnl_Client.Caption = ""
   pnl_NumSol.Caption = ""
   pnl_Produc.Caption = ""
   pnl_Modali.Caption = ""
   pnl_EjeVta.Caption = ""
   pnl_FecIng.Caption = ""
   pnl_IniEva.Caption = ""
   
   tab_Princi.Tab = 0
   tab_DatCli.Tab = 0
   tab_DatCyg.Tab = 0
   
   Call fs_LimpiaItem
End Sub

Private Sub fs_LimpiaItem()
   'Datos del Cliente
   pnl_Tit_FecNac.Caption = ""
   pnl_Tit_Paises.Caption = ""
   pnl_Tit_LugNac.Caption = ""
   pnl_Tit_EstCiv.Caption = ""
   pnl_Tit_NivEst.Caption = ""
   pnl_Tit_Direcc.Caption = ""
   pnl_Tit_Celula.Caption = ""
   pnl_Tit_Telefo.Caption = ""
   pnl_Tit_DirEle.Caption = ""
   pnl_Tit_RegCyg.Caption = ""
   pnl_Tit_Profes.Caption = ""

   pnl_Tit_OcuPri.Caption = ""
   Call gs_LimpiaGrid(grd_Tit_ActPri)
   
   pnl_Tit_OcuSec.Caption = ""
   Call gs_LimpiaGrid(grd_Tit_ActSec)

   'Datos del Cónyuge
   pnl_Cyg_DocIde.Caption = ""
   pnl_Cyg_ApeNom.Caption = ""
   pnl_Cyg_FecNac.Caption = ""
   pnl_Cyg_Paises.Caption = ""
   pnl_Cyg_LugNac.Caption = ""
   pnl_Cyg_NivEst.Caption = ""
   pnl_Cyg_Celula.Caption = ""
   pnl_Cyg_DirEle.Caption = ""
   pnl_Cyg_Profes.Caption = ""
   
   pnl_Cyg_OcuPri.Caption = ""
   Call gs_LimpiaGrid(grd_Cyg_ActPri)
   
   pnl_Cyg_OcuSec.Caption = ""
   Call gs_LimpiaGrid(grd_Cyg_ActSec)
   
   'Datos de Crédito
   pnl_Cre_TipMon.Caption = ""
   pnl_Cre_TasInt.Caption = "0.00 "
   pnl_Cre_ComVta.Caption = "0.00 "
   pnl_Cre_ApoPro.Caption = "0.00 "
   pnl_Cre_MonSol_Dol.Caption = "0.00 "
   pnl_Cre_MonSol_Sol.Caption = "0.00 "
   pnl_Cre_MonSol_MPr.Caption = "0.00 "
   pnl_Cre_MonApr_Dol.Caption = "0.00 "
   pnl_Cre_MonApr_Sol.Caption = "0.00 "
   pnl_Cre_MonApr_MPr.Caption = "0.00 "
   pnl_Cre_CuoFij_Dol.Caption = "0.00 "
   pnl_Cre_CuoIni_Dol.Caption = "0.00 "
   pnl_Cre_CuoFin_Dol.Caption = "0.00 "
   pnl_Cre_CuoFij_Sol.Caption = "0.00 "
   pnl_Cre_CuoIni_Sol.Caption = "0.00 "
   pnl_Cre_CuoFin_Sol.Caption = "0.00 "
   pnl_Cre_CuoFij_MPr.Caption = "0.00 "
   pnl_Cre_CuoIni_MPr.Caption = "0.00 "
   pnl_Cre_CuoFin_MPr.Caption = "0.00 "
   pnl_Cre_PlaApr.Caption = "0 "
   pnl_Cre_CuoExt.Caption = ""
   pnl_Cre_PerGra.Caption = "0 "
   pnl_Cre_ILDTit.Caption = "0.00 "
   pnl_Cre_ILDCyg.Caption = "0.00 "
   pnl_Cre_CuoRen.Caption = "0.00 "
   pnl_Cre_TCaDol.Caption = "0.000000 "
   pnl_Cre_TCaMPr.Caption = "0.000000 "

   'Datos de Tasación
   pnl_Tas_EmpPer.Caption = ""
   pnl_Tas_NumInf.Caption = ""
   pnl_Tas_FecEmi.Caption = ""
   pnl_Tas_FecEva.Caption = ""
   pnl_Tas_NomPer.Caption = ""
   pnl_Tas_ValCom.Caption = "0.00 "
   pnl_Tas_ValRea.Caption = "0.00 "
   pnl_Tas_AreTer.Caption = "0.00 "
   pnl_Tas_AreCon.Caption = "0.00 "
   pnl_Tas_VCoEs1.Caption = "0.00 "
   pnl_Tas_VReEs1.Caption = "0.00 "
   pnl_Tas_ATeEs1.Caption = "0.00 "
   pnl_Tas_ACoEs1.Caption = "0.00 "
   pnl_Tas_VCoEs2.Caption = "0.00 "
   pnl_Tas_VReEs2.Caption = "0.00 "
   pnl_Tas_ATeEs2.Caption = "0.00 "
   pnl_Tas_ACoEs2.Caption = "0.00 "
   pnl_Tas_VCoDep.Caption = "0.00 "
   pnl_Tas_VReDep.Caption = "0.00 "
   pnl_Tas_ATeDep.Caption = "0.00 "
   pnl_Tas_ACoDep.Caption = "0.00 "
   pnl_Tas_TotVCo.Caption = "0.00 "
   pnl_Tas_TotVRe.Caption = "0.00 "
   pnl_Tas_TotATe.Caption = "0.00 "
   pnl_Tas_TotACo.Caption = "0.00 "
   txt_Tas_Observ.Text = ""

   'Datos de Seguros
   pnl_Seg_SegPre.Caption = ""
   pnl_Seg_SegViv.Caption = ""
   pnl_Seg_InfPre.Caption = ""
   pnl_Seg_EvaPre.Caption = ""
   pnl_Seg_AplPre.Caption = ""
   pnl_Seg_FoiPre.Caption = "0.000000000 "
   pnl_Seg_PolTit.Caption = ""
   pnl_Seg_EmiTit.Caption = ""
   pnl_Seg_PolCyg.Caption = ""
   pnl_Seg_InfViv.Caption = ""
   pnl_Seg_EvaViv.Caption = ""
   pnl_Seg_AplViv.Caption = ""
   pnl_Seg_FoiViv.Caption = "0.000000000 "
   pnl_Seg_PolViv.Caption = ""
   pnl_Seg_EmiViv.Caption = ""
   txt_Seg_ObsEva.Text = ""
   txt_Seg_ObsPol.Text = ""
   
   'Datos de Legal
   txt_Leg_InfLeg.Text = ""
   pnl_Leg_AprCom.Caption = ""
   pnl_Leg_FirCon.Caption = ""
   pnl_Leg_RepLeg.Caption = ""
   pnl_Leg_Notari.Caption = ""
   pnl_Leg_FecBlq.Caption = ""
   pnl_Leg_DocReg.Caption = ""
   txt_Leg_ObsBlq.Text = ""

   'Datos de COFIDE
   pnl_Cof_NumCar.Caption = ""
   pnl_Cof_FecEmi.Caption = ""
   pnl_Cof_FecVal.Caption = ""
   pnl_Cof_NomBan.Caption = ""
   pnl_Cof_NumCta.Caption = ""
   pnl_Cof_NumOpe.Caption = ""
   pnl_Cof_TipMon.Caption = ""
   pnl_Cof_Import.Caption = "0.00 "
   pnl_Cof_TasInt.Caption = "0.00 "
   pnl_Cof_TasCom.Caption = "0.00 "
   
   'Datos de Inmueble
   pnl_Inm_Direcc.Caption = ""
   pnl_Inm_TipPro.Caption = ""
   pnl_Inm_JurEmp.Caption = ""
   pnl_Inm_JurRep.Caption = ""
   pnl_Inm_JurDir.Caption = ""
   pnl_Inm_NatTit.Caption = ""
   pnl_Inm_NatCyg.Caption = ""
   
   'Autorización de Desembolso
   pnl_Aut_FueFin.Caption = ""
   pnl_Aut_BonoBP.Caption = ""
   pnl_Aut_FecDes.Caption = ""
   txt_Aut_Observ.Text = ""
   
   'Activación de Operación
   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   Call gs_LimpiaGrid(grd_CofNCo_Listad)
   Call gs_LimpiaGrid(grd_CofCon_Listad)

   'Variables Globales
   l_str_Cli_CodCiu = ""
   l_int_Cli_TipVia = 0
   l_str_Cli_NomVia = ""
   l_str_Cli_NumVia = ""
   l_str_Cli_IntDpt = ""
   l_int_Cli_TipZon = 0
   l_str_Cli_NomZon = ""
   l_str_Cli_Refere = ""
   l_str_Cli_UbiGeo = "000000"
   l_str_Cli_Telefo = ""
   
   l_dbl_Cre_MtoPre = 0
   l_dbl_Cre_TasInt = 0
   l_dbl_Cre_TasCof = 0
   l_dbl_Cre_ComCof = 0
   l_int_Cre_NumCuo = 0
   l_int_Cre_CuoExt = 0
   l_int_Cre_PerGra = 0
   l_dbl_Cre_PreMPr = 0
   l_dbl_Cre_PreSol = 0
   l_dbl_Cre_PreDol = 0
   l_dbl_Cre_ComVta = 0
   l_dbl_Cre_ApoPro = 0
   l_dbl_Cre_TCaDol = 0
   l_dbl_Cre_TCaMPr = 0
                       
   l_str_Inm_UbiGeo = "000000"
   l_str_Inm_CodPry = ""
   l_int_Inm_PryMCs = 0
   l_int_Inm_TipVia = 0
   l_str_Inm_NomVia = ""
   l_str_Inm_NumVia = ""
   l_str_Inm_IntDpt = ""
   l_int_Inm_TipZon = 0
   l_str_Inm_NomZon = ""
   l_str_Inm_Refere = ""
   l_str_Inm_Telefo = ""
                       
   l_int_Aut_BonoBP = 0
   l_str_Aut_FecDes = "0"
   l_str_Aut_FueFin = ""
                       
   l_str_Leg_AprCom = "0"
                       
   l_dbl_Seg_FoiPre = 0
   l_dbl_Seg_FoiViv = 0
   l_int_Seg_AplPre = 0
   l_int_Seg_TipSeg = 0
   l_int_Seg_AplViv = 0
   l_str_Seg_EmpDes = ""
   l_str_Seg_EmpViv = ""
   
   l_str_Tas_EmpPer = ""
   l_str_Tas_NomPer = ""
   l_dbl_Tas_ValCom = 0
   l_dbl_Tas_ValFab = 0
   l_int_Tas_TipMon = 0
   
   l_str_Cof_NumOpe = ""
   l_int_Cof_TipMon = 0
   l_dbl_Cof_MtoDes = 0
   
   l_str_CliNCo_PriVct = "0"
   l_str_CliNCo_UltVct = "0"
   l_str_CliNCo_PrxVct = "0"
   l_dbl_CliNCo_CuoIni = 0
   l_dbl_CliNCo_CuoFin = 0
   l_dbl_CliNCo_CuoFij = 0
   
   l_str_CofNCo_PriVct = "0"
   l_str_CofNCo_UltVct = "0"
   l_str_CofNCo_PrxVct = "0"
   l_dbl_CofNCo_CuoIni = 0
   l_dbl_CofNCo_CuoFin = 0
   l_dbl_CofNCo_CuoFij = 0
                         
   l_str_CofCon_PriVct = "0"
   l_str_CofCon_UltVct = "0"
   l_str_CofCon_PrxVct = "0"
   l_dbl_CofCon_CuoIni = 0
   l_dbl_CofCon_CuoFin = 0
   l_dbl_CofCon_CuoFij = 0
   
   l_dbl_PorNCo = 0
   l_dbl_PorCon = 0
   l_dbl_ImpNCo = 0
   l_dbl_ImpCon = 0
   l_dbl_OtrCar = 0
   l_dbl_SegPre = 0
   l_dbl_SegViv = 0
   l_int_ClaCre = 0
   l_int_IndITF = 0
   l_str_NumOpe = ""

   pnl_CliNCo_Capita.Caption = "0.00 "
   pnl_CliNCo_Intere.Caption = "0.00 "
   pnl_CliNCo_SegPre.Caption = "0.00 "
   pnl_CliNCo_SegViv.Caption = "0.00 "
   pnl_CliNCo_OtrCar.Caption = "0.00 "
   pnl_CliNCo_TotCuo.Caption = "0.00 "

   pnl_CliCon_Capita.Caption = "0.00 "
   pnl_CliCon_Intere.Caption = "0.00 "
   pnl_CliCon_TotCuo.Caption = "0.00 "

   pnl_CofCon_Capita.Caption = "0.00 "
   pnl_CofCon_Intere.Caption = "0.00 "
   pnl_CofCon_Comisi.Caption = "0.00 "
   pnl_CofCon_TotCuo.Caption = "0.00 "
   
   pnl_CofNCo_Capita.Caption = "0.00 "
   pnl_CofNCo_Intere.Caption = "0.00 "
   pnl_CofNCo_Comisi.Caption = "0.00 "
   pnl_CofNCo_TotCuo.Caption = "0.00 "
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumSol.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   tab_Princi.Enabled = Not p_Habilita
   tab_Cronog.Enabled = Not p_Habilita
End Sub

Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   cmd_Grabar.Enabled = p_Habilita
End Sub

Private Sub fs_Buscar_DatGen()
   g_rst_Princi.MoveFirst
   
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO)
   moddat_g_str_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))


   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)
   
   'Ejecutivo de Ventas
   moddat_g_str_CodEje = Trim(g_rst_Princi!SOLMAE_EJEVTA)
   moddat_g_str_EjeVta = moddat_gf_Buscar_NomEje(moddat_g_str_CodEje)

   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS

   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))

   'Fecha de Ingreso
   moddat_g_str_FecIng = Right(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 2) & "/" & Mid(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 5, 2) & "/" & Left(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 4)
   
   'Información de Seguros
   pnl_Seg_SegPre.Caption = Trim(moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES)) & " / " & Trim(moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG))
   pnl_Seg_SegViv.Caption = Trim(moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGVIV))
   
   
   l_str_Seg_EmpDes = g_rst_Princi!SOLMAE_ESGDES
   l_int_Seg_TipSeg = g_rst_Princi!SOLMAE_TIPSEG
   l_str_Seg_EmpViv = g_rst_Princi!SOLMAE_ESGDES
End Sub

Private Sub fs_Buscar_SegDet()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_AutDes) & " "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_str_FecOcu = Right(CStr(g_rst_Princi!SEGDET_FECOCU), 2) & "/" & Mid(CStr(g_rst_Princi!SEGDET_FECOCU), 5, 2) & "/" & Left(CStr(g_rst_Princi!SEGDET_FECOCU), 4)
      
      Select Case g_rst_Princi!SEGDET_CODOCU
         Case 11:    l_str_IniEva = r_str_FecOcu
      End Select
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If Len(Trim(l_str_IniEva)) > 0 Then
      pnl_IniEva.Caption = l_str_IniEva
   End If
End Sub

Private Sub fs_Buscar_InfSol()
   Call fs_Buscar_DatCli
   Call fs_Buscar_DatCyg
   Call fs_Buscar_DatCre
   Call fs_Buscar_DatTas
   Call fs_Buscar_DatSeg
   Call fs_Buscar_DatLeg
   
   If moddat_g_str_CodPrd = "001" Then
      Call fs_Buscar_DatCof
   End If
   
   Call fs_Buscar_DatInm
   Call fs_Buscar_DatAut
End Sub

Private Sub fs_Buscar_DatCli()
   Dim r_str_Depart     As String
   Dim r_str_Provin     As String
   Dim r_str_Distri     As String
   Dim r_str_TipVia     As String
   Dim r_str_TipZon     As String

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   g_rst_Princi.MoveFirst

   pnl_Tit_FecNac.Caption = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
   pnl_Tit_Celula.Caption = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
   pnl_Tit_Telefo.Caption = Trim(g_rst_Princi!DatGen_Telefo & "")
   pnl_Tit_DirEle.Caption = Trim(g_rst_Princi!DatGen_DirEle & "")

   pnl_Tit_EstCiv.Caption = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV))
   
   If g_rst_Princi!DatGen_RegCyg > 0 Then
      pnl_Tit_RegCyg.Caption = moddat_gf_Consulta_ParDes("206", CStr(g_rst_Princi!DatGen_RegCyg))
   End If
   
   pnl_Tit_NivEst.Caption = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

   'País de Nacimiento
   pnl_Tit_Paises.Caption = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))

   'Profesión
   pnl_Tit_Profes.Caption = moddat_gf_Consulta_ParDes("501", Trim(g_rst_Princi!DatGen_Profes))

   r_str_Depart = ""
   r_str_Provin = ""
   r_str_Distri = ""
   
   If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
      'Departamento
      r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000")
      
      'Provincia
      r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00")
      
      'Distrito
      r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      
      pnl_Tit_LugNac.Caption = r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   End If

   r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia))
   r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon))

   pnl_Tit_Direcc.Caption = r_str_TipVia & " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero)
   
   If Len(Trim(Trim(g_rst_Princi!DatGen_IntDpt))) > 0 Then
      pnl_Tit_Direcc.Caption = pnl_Tit_Direcc.Caption & " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")"
   End If
   
   If Len(Trim(Trim(g_rst_Princi!DatGen_NomZon))) > 0 Then
      pnl_Tit_Direcc.Caption = pnl_Tit_Direcc.Caption & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!DatGen_NomZon) & Chr(13) & Chr(10)
   Else
      pnl_Tit_Direcc.Caption = pnl_Tit_Direcc.Caption & Chr(13) & Chr(10)
   End If
   
   'Departamento
   r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000")
   
   'Provincia
   r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00")
   
   'Distrito
   r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
   
   pnl_Tit_Direcc.Caption = pnl_Tit_Direcc.Caption & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   If g_rst_Princi!DATGEN_CYGTDO > 0 Then
      moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
      moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO)
   End If
   
   l_str_Cli_CodCiu = g_rst_Princi!DATGEN_CODCIU
   
   l_int_Cli_TipVia = g_rst_Princi!DatGen_TipVia
   l_str_Cli_NomVia = Trim(g_rst_Princi!DatGen_NomVia & "")
   l_str_Cli_NumVia = Trim(g_rst_Princi!DatGen_Numero & "")
   l_str_Cli_IntDpt = Trim(g_rst_Princi!DatGen_IntDpt & "")
   l_int_Cli_TipZon = g_rst_Princi!DatGen_TipZon
   l_str_Cli_NomZon = Trim(g_rst_Princi!DatGen_NomZon & "")
   l_str_Cli_Refere = Trim(g_rst_Princi!DatGen_Refere & "")
   l_str_Cli_UbiGeo = Trim(g_rst_Princi!DatGen_Ubigeo & "")
   l_str_Cli_Telefo = Trim(g_rst_Princi!DatGen_Telefo & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Buscar_ActEco(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1, grd_Tit_ActPri, pnl_Tit_OcuPri)
   Call fs_Buscar_ActEco(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2, grd_Tit_ActSec, pnl_Tit_OcuSec)
End Sub

Private Sub fs_Buscar_DatCyg()
   Dim r_str_Depart     As String
   Dim r_str_Provin     As String
   Dim r_str_Distri     As String
   Dim r_str_TipVia     As String
   Dim r_str_TipZon     As String

   If moddat_g_int_CygTDo = 0 Then
      Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_CygTDo) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_CygNDo & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   g_rst_Princi.MoveFirst

   pnl_Cyg_DocIde.Caption = CStr(g_rst_Princi!DatGen_TipDoc) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
   pnl_Cyg_ApeNom.Caption = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & " " & Trim(g_rst_Princi!DatGen_Nombre)

   pnl_Cyg_FecNac.Caption = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
   pnl_Cyg_Celula.Caption = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
   pnl_Cyg_DirEle.Caption = Trim(g_rst_Princi!DatGen_DirEle & "")

   pnl_Cyg_NivEst.Caption = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

   'País de Nacimiento
   pnl_Cyg_Paises.Caption = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))

   'Profesión
   pnl_Cyg_Profes.Caption = moddat_gf_Consulta_ParDes("501", Trim(g_rst_Princi!DatGen_Profes))

   r_str_Depart = ""
   r_str_Provin = ""
   r_str_Distri = ""
   
   If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
      'Departamento
      r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000")
      
      'Provincia
      r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00")
      
      'Distrito
      r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      
      pnl_Cyg_LugNac.Caption = r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Buscar_ActEco(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1, grd_Cyg_ActPri, pnl_Cyg_OcuPri)
   Call fs_Buscar_ActEco(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2, grd_Cyg_ActSec, pnl_Cyg_OcuSec)
End Sub

Private Sub fs_Buscar_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer, p_Listad As MSFlexGrid, p_NomOcu As SSPanel)
   Dim r_str_Depart     As String
   Dim r_str_Provin     As String
   Dim r_str_Distri     As String
   Dim r_str_TipVia     As String
   Dim r_str_TipZon     As String
   Dim r_str_TipDoc     As String
   Dim l_rst_Genera     As ADODB.Recordset
   
   Call gs_LimpiaGrid(p_Listad)
   p_NomOcu.Caption = ""
   
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Ocupación
   p_NomOcu.Caption = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ActEco_CodAct))
   
   Select Case g_rst_Princi!ActEco_CodAct
      Case 11, 31, 41
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_TipDoc) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_NumDoc) & "' "
      
         If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
            Exit Sub
         End If
   
         l_rst_Genera.MoveFirst
         
         'Documento de Identidad
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Documento de Identidad"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
      
         'Razón Social
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Razón Social"
      
         p_Listad.Col = 1
         p_Listad.Text = Trim(l_rst_Genera!DATGEN_RAZSOC)
      
         'Nombre Comercial
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Nombre Comercial"
      
         p_Listad.Col = 1
         p_Listad.Text = Trim(l_rst_Genera!DATGEN_NOMCOM & "")
      
         'Giro Comercial
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Giro Comercial"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Busca_GirCom(Trim(l_rst_Genera!DATGEN_GCOMCO))
      
         If Len(Trim(l_rst_Genera!DATGEN_GCOMNO & "")) > 0 Then
            p_Listad.Text = p_Listad.Text & " - " & Trim(l_rst_Genera!DATGEN_GCOMNO)
         End If
      
         'Dirección
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Dirección Empresa"
      
         p_Listad.Col = 1
         r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(l_rst_Genera!DatGen_TipVia))
         r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(l_rst_Genera!DatGen_TipZon))

         p_Listad.Text = r_str_TipVia & " " & Trim(l_rst_Genera!DatGen_NomVia & "") & " " & Trim(l_rst_Genera!DatGen_Numero & "")

         If Len(Trim(Trim(l_rst_Genera!DatGen_IntDpt & ""))) > 0 Then
            p_Listad.Text = p_Listad.Text & " (" & Trim(l_rst_Genera!DatGen_IntDpt) & ")"
         End If

         If Len(Trim(Trim(l_rst_Genera!DatGen_NomZon & ""))) > 0 Then
            p_Listad.Text = p_Listad.Text & " - " & r_str_TipZon & " " & Trim(l_rst_Genera!DatGen_NomZon) & " / "
         Else
            p_Listad.Text = p_Listad.Text & " / "
         End If
         
         r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 2) & "0000")
         r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 4) & "00")
         r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(l_rst_Genera!DatGen_Ubigeo))
   
         p_Listad.Text = p_Listad.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
         'Teléfono
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Teléfono(s) Empresa"
      
         p_Listad.Col = 1
         p_Listad.Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
         
         If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
            p_Listad.Text = p_Listad.Text & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         'Sucursal
         If Len(Trim(g_rst_Princi!ActEco_Sucurs & "")) > 0 Then
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Sucursal"
         
            p_Listad.Col = 1
            p_Listad.Text = Trim(g_rst_Princi!ACTECO_DEP_SUCURS & "")
            
            'Dirección Sucursal
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Dirección Sucursal"
         
            p_Listad.Col = 1
            
            r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_TipVia))
            r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_TipZon))

            p_Listad.Text = r_str_TipVia & " " & Trim(g_rst_Princi!ActEco_NomVia & "") & " " & Trim(g_rst_Princi!ActEco_Numero & "")
   
            If Len(Trim(Trim(g_rst_Princi!ActEco_IntDpt & ""))) > 0 Then
               p_Listad.Text = p_Listad.Text & " (" & Trim(g_rst_Princi!ActEco_IntDpt) & ")"
            End If
   
            If Len(Trim(Trim(g_rst_Princi!ActEco_NomZon & ""))) > 0 Then
               p_Listad.Text = p_Listad.Text & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!ActEco_NomZon) & " / "
            Else
               p_Listad.Text = p_Listad.Text & " / "
            End If
            
            r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 2) & "0000")
            r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 4) & "00")
            r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ubigeo))
      
            p_Listad.Text = p_Listad.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
            'Teléfono Sucursal
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Teléfono(s) Sucursal"
         
            p_Listad.Col = 1
            p_Listad.Text = Trim(g_rst_Princi!ActEco_Telef1 & "")
            
            If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
               p_Listad.Text = p_Listad.Text & Trim(g_rst_Princi!ActEco_Telef2 & "")
            End If
         End If
         
         If g_rst_Princi!ActEco_CodAct = 11 Then
            'Teléfono y Anexo RR.HH
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Teléfono RR.HH"
         
            p_Listad.Col = 1
            
            If Len(Trim(l_rst_Genera!DATGEN_TELERH & "")) = 0 Then
               p_Listad.Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
            Else
               p_Listad.Text = Trim(l_rst_Genera!DATGEN_TELERH & "")
            End If
            
            If Len(Trim(l_rst_Genera!DATGEN_ANEXRH & "")) > 0 Then
               p_Listad.Text = p_Listad.Text & " - " & Trim(l_rst_Genera!DATGEN_ANEXRH & "")
            End If
         
            'Cargo
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Cargo"
         
            p_Listad.Col = 1
            If Len(Trim(g_rst_Princi!ActEco_Dep_CargoN & "")) > 0 Then
               p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_CargoN)
            Else
               p_Listad.Text = moddat_gf_Consulta_ParDes("503", Trim(g_rst_Princi!ActEco_Dep_CargoC))
            End If
         
            'Area
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Area"
         
            p_Listad.Col = 1
            p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_NomAre)
            
            'Número Anexo
            If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
               p_Listad.Rows = p_Listad.Rows + 1
               p_Listad.Row = p_Listad.Rows - 1
               
               p_Listad.Col = 0
               p_Listad.Text = "Anexo"
            
               p_Listad.Col = 1
               p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx)
            End If
            
            'Teléfono Directo
            If Len(Trim(g_rst_Princi!ActEco_Dep_TelDir & "")) > 0 Then
               p_Listad.Rows = p_Listad.Rows + 1
               p_Listad.Row = p_Listad.Rows - 1
               
               p_Listad.Col = 0
               p_Listad.Text = "Teléfono Directo"
            
               p_Listad.Col = 1
               p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_TelDir)
            End If
         
            'Celular Laboral
            If Len(Trim(g_rst_Princi!ActEco_Dep_Celula)) > 0 Then
               p_Listad.Rows = p_Listad.Rows + 1
               p_Listad.Row = p_Listad.Rows - 1
               
               p_Listad.Col = 0
               p_Listad.Text = "Celular Laboral"
            
               p_Listad.Col = 1
               p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_Celula)
            End If
         
            'E-mail
            If Len(Trim(g_rst_Princi!ActEco_Dep_DirEle)) > 0 Then
               p_Listad.Rows = p_Listad.Rows + 1
               p_Listad.Row = p_Listad.Rows - 1
               
               p_Listad.Col = 0
               p_Listad.Text = "E-mail"
            
               p_Listad.Col = 1
               p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_DirEle)
            End If
         End If
         
         l_rst_Genera.Close
         Set l_rst_Genera = Nothing
         
         Call gs_UbiIniGrid(p_Listad)
         
      Case 21
         'Documento de Identidad
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Documento de Identidad"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
         
         'Dirección Tributaria
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Dirección Tributaria"
      
         p_Listad.Col = 1
         
         r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_TipVia))
         r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_TipZon))

         p_Listad.Text = r_str_TipVia & " " & Trim(g_rst_Princi!ActEco_NomVia & "") & " " & Trim(g_rst_Princi!ActEco_Numero & "")

         If Len(Trim(Trim(g_rst_Princi!ActEco_IntDpt & ""))) > 0 Then
            p_Listad.Text = p_Listad.Text & " (" & Trim(g_rst_Princi!ActEco_IntDpt) & ")"
         End If

         If Len(Trim(Trim(g_rst_Princi!ActEco_NomZon & ""))) > 0 Then
            p_Listad.Text = p_Listad.Text & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!ActEco_NomZon) & " / "
         Else
            p_Listad.Text = p_Listad.Text & " / "
         End If
         
         r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 2) & "0000")
         r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 4) & "00")
         r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ubigeo))
   
         p_Listad.Text = p_Listad.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
      
         'Teléfono
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Teléfono(s) "
      
         p_Listad.Col = 1
         p_Listad.Text = Trim(g_rst_Princi!ActEco_Telef1 & "")
         
         If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
            p_Listad.Text = p_Listad.Text & Trim(g_rst_Princi!ActEco_Telef2 & "")
         End If
         
         'Giro Comercial
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Giro Comercial"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Busca_GirCom(Trim(g_rst_Princi!ActEco_GiroCd))
      
         If Len(Trim(g_rst_Princi!ActEco_GiroNm & "")) > 0 Then
            p_Listad.Text = p_Listad.Text & " - " & Trim(g_rst_Princi!ActEco_GiroNm)
         End If
         
         'Contrato de Locación de Servicios
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Contrato Locación "
         
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
         
         If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
            g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
            g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TDoEmp) & " AND "
            g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_Ind_NDoEmp) & "' "
      
            If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
               Exit Sub
            End If
   
            l_rst_Genera.MoveFirst
         
            'Documento de Identidad
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
         
            p_Listad.Col = 0
            p_Listad.Text = "Documento Ident. Empresa"
      
            p_Listad.Col = 1
            p_Listad.Text = moddat_gf_Consulta_ParDes("203", CStr(l_rst_Genera!DatGen_EMPTDO)) & " - " & Trim(l_rst_Genera!DatGen_EMPNDO)
      
            'Razón Social
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
         
            p_Listad.Col = 0
            p_Listad.Text = "Razón Social Empresa"
      
            p_Listad.Col = 1
            p_Listad.Text = Trim(l_rst_Genera!DATGEN_RAZSOC)
         
            'Nombre Comercial
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Nombre Comercial Empresa"
         
            p_Listad.Col = 1
            p_Listad.Text = Trim(l_rst_Genera!DATGEN_NOMCOM)
         
            'Giro Comercial
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Giro Comercial Empresa"
         
            p_Listad.Col = 1
            p_Listad.Text = moddat_gf_Busca_GirCom(Trim(l_rst_Genera!DATGEN_GCOMCO))
         
            If Len(Trim(l_rst_Genera!DATGEN_GCOMNO & "")) > 0 Then
               p_Listad.Text = p_Listad.Text & " - " & Trim(l_rst_Genera!DATGEN_GCOMNO)
            End If
         
            'Dirección
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Dirección Empresa"
         
            p_Listad.Col = 1
            r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(l_rst_Genera!DatGen_TipVia))
            r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(l_rst_Genera!DatGen_TipZon))
   
            p_Listad.Text = r_str_TipVia & " " & Trim(l_rst_Genera!DatGen_NomVia & "") & " " & Trim(l_rst_Genera!DatGen_Numero & "")
   
            If Len(Trim(Trim(l_rst_Genera!DatGen_IntDpt & ""))) > 0 Then
               p_Listad.Text = p_Listad.Text & " (" & Trim(l_rst_Genera!DatGen_IntDpt) & ")"
            End If
   
            If Len(Trim(Trim(l_rst_Genera!DatGen_NomZon & ""))) > 0 Then
               p_Listad.Text = p_Listad.Text & " - " & r_str_TipZon & " " & Trim(l_rst_Genera!DatGen_NomZon) & " / "
            Else
               p_Listad.Text = p_Listad.Text & " / "
            End If
            
            r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 2) & "0000")
            r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 4) & "00")
            r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(l_rst_Genera!DatGen_Ubigeo))
      
            p_Listad.Text = p_Listad.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
            'Teléfonos
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
         
            p_Listad.Col = 0
            p_Listad.Text = "Teléfonos Empresa"
            
            p_Listad.Col = 1
            p_Listad.Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
         
            If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
               p_Listad.Text = p_Listad.Text & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
            End If
         End If
         
         Call gs_UbiIniGrid(p_Listad)
         
      Case 51
         'Documento de Identidad
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Documento de Identidad"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
         
         'Giro Comercial
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Giro Comercial"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Busca_GirCom(Trim(g_rst_Princi!ActEco_GiroCd))
      
         If Len(Trim(g_rst_Princi!ActEco_GiroNm & "")) > 0 Then
            p_Listad.Text = p_Listad.Text & " - " & Trim(g_rst_Princi!ActEco_GiroNm)
         End If
         
         Call gs_UbiIniGrid(p_Listad)
   End Select
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatCre()
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst

   pnl_Cre_TipMon.Caption = moddat_gf_Consulta_ParDes("204", g_rst_Princi!SOLMAE_TIPMON)
   pnl_Cre_TasInt.Caption = Format(g_rst_Princi!SOLMAE_TASINT, "##0.00") & " "
   pnl_Cre_ComVta.Caption = Format(g_rst_Princi!SOLMAE_COMVTA, "###,###,##0.00") & " "
   pnl_Cre_ApoPro.Caption = Format(g_rst_Princi!SOLMAE_APOPRO, "###,###,##0.00") & " "
   pnl_Cre_MonSol_Dol.Caption = Format(g_rst_Princi!SOLMAE_MTOSOL, "###,###,##0.00") & " "
   pnl_Cre_MonSol_Sol.Caption = Format(g_rst_Princi!SOLMAE_PRESOL, "###,###,##0.00") & " "
   pnl_Cre_MonSol_MPr.Caption = Format(g_rst_Princi!SOLMAE_PREMPR, "###,###,##0.00") & " "
   pnl_Cre_MonApr_Dol.Caption = Format(g_rst_Princi!SOLMAE_APRDOL, "###,###,##0.00") & " "
   pnl_Cre_MonApr_Sol.Caption = Format(g_rst_Princi!SOLMAE_APRSOL, "###,###,##0.00") & " "
   pnl_Cre_MonApr_MPr.Caption = Format(g_rst_Princi!SOLMAE_APRMPR, "###,###,##0.00") & " "
   
   pnl_Cre_TCaDol.Caption = Format(g_rst_Princi!SOLMAE_APRTCD, "###,##0.000000") & " "
   pnl_Cre_TCaMPr.Caption = Format(g_rst_Princi!SOLMAE_APRTCM, "###,##0.000000") & " "
   
   pnl_Cre_CuoFij_Dol.Caption = Format(g_rst_Princi!SOLMAE_APRCUO, "###,###,##0.00") & " "
   pnl_Cre_CuoIni_Dol.Caption = Format(g_rst_Princi!SOLMAE_APRCIN, "###,###,##0.00") & " "
   pnl_Cre_CuoFin_Dol.Caption = Format(g_rst_Princi!SOLMAE_APRCFN, "###,###,##0.00") & " "
   
   pnl_Cre_CuoFij_Sol.Caption = Format(g_rst_Princi!SOLMAE_APRCUO * CDbl(pnl_Cre_TCaDol.Caption), "###,###,##0.00") & " "
   pnl_Cre_CuoIni_Sol.Caption = Format(g_rst_Princi!SOLMAE_APRCIN * CDbl(pnl_Cre_TCaDol.Caption), "###,###,##0.00") & " "
   pnl_Cre_CuoFin_Sol.Caption = Format(g_rst_Princi!SOLMAE_APRCFN * CDbl(pnl_Cre_TCaDol.Caption), "###,###,##0.00") & " "
   
   pnl_Cre_CuoFij_MPr.Caption = Format(CDbl(pnl_Cre_CuoFij_Sol.Caption) / CDbl(pnl_Cre_TCaMPr.Caption), "###,###,##0.00") & " "
   pnl_Cre_CuoIni_MPr.Caption = Format(CDbl(pnl_Cre_CuoIni_Sol.Caption) / CDbl(pnl_Cre_TCaMPr.Caption), "###,###,##0.00") & " "
   pnl_Cre_CuoFin_MPr.Caption = Format(CDbl(pnl_Cre_CuoFin_Sol.Caption) / CDbl(pnl_Cre_TCaMPr.Caption), "###,###,##0.00") & " "
   
   pnl_Cre_PlaApr.Caption = Format(g_rst_Princi!SOLMAE_APRPLA, "##0") & " "
   pnl_Cre_PerGra.Caption = Format(g_rst_Princi!SOLMAE_APRPGR, "##0") & " "
   pnl_Cre_CuoExt.Caption = moddat_gf_Consulta_ParDes("223", g_rst_Princi!SOLMAE_CUOANO)
   pnl_Cre_ILDTit.Caption = Format(g_rst_Princi!SOLMAE_APRIN1 + g_rst_Princi!SOLMAE_APRIN2, "###,###,##0.00") & " "
   pnl_Cre_ILDCyg.Caption = Format(g_rst_Princi!SOLMAE_APRIN3 + g_rst_Princi!SOLMAE_APRIN4, "###,###,##0.00") & " "
   pnl_Cre_CuoRen.Caption = Format(g_rst_Princi!SOLMAE_APRRCR, "##0.00") & " "
   
   
   
   pnl_Cof_TasInt.Caption = Format(g_rst_Princi!SOLMAE_INTCOF, "##0.00") & " "
   pnl_Cof_TasCom.Caption = Format(g_rst_Princi!SOLMAE_COMCOF, "##0.00") & " "
   
   
   'Variables
   l_dbl_Cre_MtoPre = g_rst_Princi!SOLMAE_APRMPR
   l_dbl_Cre_PreSol = g_rst_Princi!SOLMAE_APRSOL
   l_dbl_Cre_PreDol = g_rst_Princi!SOLMAE_APRDOL
   l_dbl_Cre_TasInt = g_rst_Princi!SOLMAE_TASINT
   l_dbl_Cre_TasCof = g_rst_Princi!SOLMAE_INTCOF
   l_dbl_Cre_ComCof = g_rst_Princi!SOLMAE_COMCOF
   l_int_Cre_NumCuo = g_rst_Princi!SOLMAE_APRPLA
   l_int_Cre_CuoExt = g_rst_Princi!SOLMAE_CUOANO
   l_int_Cre_PerGra = g_rst_Princi!SOLMAE_APRPGR
   l_dbl_Cre_ComVta = g_rst_Princi!SOLMAE_COMVTA
   l_dbl_Cre_ApoPro = g_rst_Princi!SOLMAE_APOPRO
   l_dbl_Cre_TCaDol = g_rst_Princi!SOLMAE_APRTCD
   l_dbl_Cre_TCaMPr = g_rst_Princi!SOLMAE_APRTCM
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatTas()
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Empresa de Peritaje
   pnl_Tas_EmpPer.Caption = moddat_gf_Consulta_ParDes("507", Format(g_rst_Princi!EVATAS_CODEMP, "000000"))

   pnl_Tas_NumInf.Caption = Trim(g_rst_Princi!EVATAS_NUMINF)
   pnl_Tas_FecEmi.Caption = Right(CStr(g_rst_Princi!EVATAS_FECEMI), 2) & "/" & Mid(CStr(g_rst_Princi!EVATAS_FECEMI), 5, 2) & "/" & Left(CStr(g_rst_Princi!EVATAS_FECEMI), 4)
   pnl_Tas_FecEva.Caption = Right(CStr(g_rst_Princi!EVATAS_FECEVA), 2) & "/" & Mid(CStr(g_rst_Princi!EVATAS_FECEVA), 5, 2) & "/" & Left(CStr(g_rst_Princi!EVATAS_FECEVA), 4)
   pnl_Tas_NomPer.Caption = Trim(g_rst_Princi!EVATAS_NOMPER)
   
   pnl_Tas_ValCom.Caption = Format(g_rst_Princi!EVATAS_VALCOM, "###,###,##0.00") & " "
   pnl_Tas_ValRea.Caption = Format(g_rst_Princi!EVATAS_VALFAB, "###,###,##0.00") & " "
   pnl_Tas_AreTer.Caption = Format(g_rst_Princi!EVATAS_ARETER, "###,###,##0.00") & " "
   pnl_Tas_AreCon.Caption = Format(g_rst_Princi!EVATAS_ARECON, "###,###,##0.00") & " "
   
   pnl_Tas_VCoEs1.Caption = Format(g_rst_Princi!EVATAS_VCOES1, "###,###,##0.00") & " "
   pnl_Tas_VReEs1.Caption = Format(g_rst_Princi!EVATAS_VREES1, "###,###,##0.00") & " "
   pnl_Tas_ATeEs1.Caption = Format(g_rst_Princi!EVATAS_ATEES1, "###,###,##0.00") & " "
   pnl_Tas_ACoEs1.Caption = Format(g_rst_Princi!EVATAS_ACOES1, "###,###,##0.00") & " "
   
   pnl_Tas_VCoEs2.Caption = Format(g_rst_Princi!EVATAS_VCOES2, "###,###,##0.00") & " "
   pnl_Tas_VReEs2.Caption = Format(g_rst_Princi!EVATAS_VREES2, "###,###,##0.00") & " "
   pnl_Tas_ATeEs2.Caption = Format(g_rst_Princi!EVATAS_ATEES2, "###,###,##0.00") & " "
   pnl_Tas_ACoEs2.Caption = Format(g_rst_Princi!EVATAS_ACOES2, "###,###,##0.00") & " "
   
   pnl_Tas_VCoDep.Caption = Format(g_rst_Princi!EVATAS_VCODEP, "###,###,##0.00") & " "
   pnl_Tas_VReDep.Caption = Format(g_rst_Princi!EVATAS_VREDEP, "###,###,##0.00") & " "
   pnl_Tas_ATeDep.Caption = Format(g_rst_Princi!EVATAS_ATEDEP, "###,###,##0.00") & " "
   pnl_Tas_ACoDep.Caption = Format(g_rst_Princi!EVATAS_ACODEP, "###,###,##0.00") & " "
   
   pnl_Tas_TotVCo.Caption = Format(g_rst_Princi!EVATAS_VALCOM + g_rst_Princi!EVATAS_VCOES1 + g_rst_Princi!EVATAS_VCOES2 + g_rst_Princi!EVATAS_VCODEP, "###,###,##0.00") & " "
   pnl_Tas_TotVRe.Caption = Format(g_rst_Princi!EVATAS_VALFAB + g_rst_Princi!EVATAS_VREES1 + g_rst_Princi!EVATAS_VREES2 + g_rst_Princi!EVATAS_VREDEP, "###,###,##0.00") & " "
   pnl_Tas_TotATe.Caption = Format(g_rst_Princi!EVATAS_ARETER + g_rst_Princi!EVATAS_ATEES1 + g_rst_Princi!EVATAS_ATEES2 + g_rst_Princi!EVATAS_ATEDEP, "###,###,##0.00") & " "
   pnl_Tas_TotACo.Caption = Format(g_rst_Princi!EVATAS_ARECON + g_rst_Princi!EVATAS_ACOES1 + g_rst_Princi!EVATAS_ACOES2 + g_rst_Princi!EVATAS_ACODEP, "###,###,##0.00") & " "
   
   txt_Tas_Observ.Text = Trim(g_rst_Princi!EVATAS_OBSERV & "")
   
   l_str_Tas_EmpPer = Format(g_rst_Princi!EVATAS_CODEMP, "000000")
   l_str_Tas_NomPer = Trim(g_rst_Princi!EVATAS_NOMPER)
   l_dbl_Tas_ValCom = CDbl(Format(g_rst_Princi!EVATAS_VALCOM + g_rst_Princi!EVATAS_VCOES1 + g_rst_Princi!EVATAS_VCOES2 + g_rst_Princi!EVATAS_VCODEP, "###,###,##0.00"))
   l_dbl_Tas_ValFab = CDbl(Format(g_rst_Princi!EVATAS_VALFAB + g_rst_Princi!EVATAS_VREES1 + g_rst_Princi!EVATAS_VREES2 + g_rst_Princi!EVATAS_VREDEP, "###,###,##0.00"))
   l_int_Tas_TipMon = g_rst_Princi!EVATAS_TIPMON
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_Seg_ObsEva_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Seg_ObsPol_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Tas_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_Buscar_DatLeg()
   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   txt_Leg_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLEG)
   
   If g_rst_Princi!EVALEG_APRCOM > 0 Then
      pnl_Leg_AprCom.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_APRCOM))
      
      l_str_Leg_AprCom = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_APRCOM))
   End If
   
   If g_rst_Princi!EVALEG_FIRCON > 0 Then
      pnl_Leg_FirCon.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
      pnl_Leg_RepLeg.Caption = Trim(g_rst_Princi!EVALEG_REPLG1)
      
      If Len(Trim(g_rst_Princi!EVALEG_REPLG2)) > 0 Then
         pnl_Leg_RepLeg.Caption = pnl_Leg_RepLeg.Caption & " / " & Trim(g_rst_Princi!EVALEG_REPLG2)
      End If
      pnl_Leg_Notari.Caption = moddat_gf_Consulta_ParDes("509", Trim(g_rst_Princi!EVALEG_BLQNOT))
   End If
   
   If g_rst_Princi!EVALEG_BLQFEC > 0 Then
      pnl_Leg_FecBlq.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_BLQFEC))
      
      If g_rst_Princi!EVALEG_TIPDOC = 1 Or g_rst_Princi!EVALEG_TIPDOC = 2 Then
         pnl_Leg_DocReg.Caption = Trim(moddat_gf_Consulta_ParDes("026", CStr(g_rst_Princi!EVALEG_TIPDOC)))
         pnl_Leg_DocReg.Caption = pnl_Leg_DocReg.Caption & " NRO.: " & Trim(g_rst_Princi!EVALEG_PARFIC) & " - ASIENTO: " & Trim(g_rst_Princi!EVALEG_NUMASI)
      Else
         pnl_Leg_DocReg.Caption = "TOMO: " & Trim(g_rst_Princi!EVALEG_BLQTOM) & " - " & "FOJAS: " & Trim(g_rst_Princi!EVALEG_BLQFOJ) & " - " & "LIBRO: " & Trim(g_rst_Princi!EVALEG_BLQLIB)
      End If
      
      txt_Leg_ObsBlq.Text = Trim(g_rst_Princi!EVALEG_BLQOBS & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatSeg()
   g_str_Parame = "SELECT * FROM TRA_EVASEG WHERE "
   g_str_Parame = g_str_Parame & "EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
     
   'Cargar Datos de Evaluación
   pnl_Seg_InfPre.Caption = Trim(g_rst_Princi!EVASEG_INFPRE & "")
   pnl_Seg_EvaPre.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAPRE))
   pnl_Seg_AplPre.Caption = moddat_gf_Consulta_ParDes("227", g_rst_Princi!EVASEG_TAPPRE)
   pnl_Seg_FoiPre.Caption = Format(g_rst_Princi!EVASEG_TASPRE, "###,###,##0.000000000") & " "
   
   pnl_Seg_InfViv.Caption = Trim(g_rst_Princi!EVASEG_INFVIV & "")
   pnl_Seg_EvaViv.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAVIV))
   pnl_Seg_AplViv.Caption = moddat_gf_Consulta_ParDes("227", g_rst_Princi!EVASEG_TAPVIV)
   pnl_Seg_FoiViv.Caption = Format(g_rst_Princi!EVASEG_TASVIV, "###,###,##0.000000000") & " "
   
   txt_Seg_ObsEva.Text = Trim(g_rst_Princi!EVASEG_OBSERV & "")
   
   l_int_Seg_AplPre = g_rst_Princi!EVASEG_TAPPRE
   l_dbl_Seg_FoiPre = g_rst_Princi!EVASEG_TASPRE
   
   l_int_Seg_AplViv = g_rst_Princi!EVASEG_TAPVIV
   l_dbl_Seg_FoiViv = g_rst_Princi!EVASEG_TASVIV
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Obteniendo Información de Póliza de Seguro
   g_str_Parame = "SELECT * FROM TRA_POLIZA WHERE "
   g_str_Parame = g_str_Parame & "POLIZA_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   pnl_Seg_PolTit.Caption = Trim(g_rst_Princi!POLIZA_NUMDES & "")
   pnl_Seg_PolCyg.Caption = Trim(g_rst_Princi!POLIZA_NUMCYG & "")
   pnl_Seg_EmiTit.Caption = gf_FormatoFecha(CStr(g_rst_Princi!POLIZA_FEMDES))

   pnl_Seg_PolViv.Caption = Trim(g_rst_Princi!POLIZA_NUMVIV & "")
   pnl_Seg_EmiViv.Caption = gf_FormatoFecha(CStr(g_rst_Princi!POLIZA_FEMVIV))
   
   txt_Seg_ObsPol.Text = Trim(g_rst_Princi!POLIZA_OBSERV & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatCof()
   g_str_Parame = "SELECT * FROM TRA_DETCOF WHERE "
   g_str_Parame = g_str_Parame & "DETCOF_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   pnl_Cof_NumCar.Caption = Trim(g_rst_Princi!DETCOF_NUMCAR)
   pnl_Cof_NumOpe.Caption = Trim(g_rst_Princi!DETCOF_NUMOPE)
   pnl_Cof_Import.Caption = Format(g_rst_Princi!DETCOF_IMPORT, "###,###,#0.00") & " "
      
   l_str_Cof_NumOpe = Trim(g_rst_Princi!DETCOF_NUMOPE)
   l_dbl_Cof_MtoDes = g_rst_Princi!DETCOF_IMPORT
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   g_str_Parame = "SELECT * FROM TRA_CARCOF WHERE "
   g_str_Parame = g_str_Parame & "CARCOF_NUMCAR = '" & pnl_Cof_NumCar.Caption & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
      
   pnl_Cof_FecEmi.Caption = gf_FormatoFecha(CStr(g_rst_Princi!CARCOF_FECEMI))
   pnl_Cof_FecVal.Caption = gf_FormatoFecha(CStr(g_rst_Princi!CARCOF_FECVAL))
   pnl_Cof_NomBan.Caption = moddat_gf_Consulta_ParDes("505", g_rst_Princi!CARCOF_CODBAN)
   pnl_Cof_NumCta.Caption = Trim(g_rst_Princi!CARCOF_NUMCTA & "")
   
   pnl_Cof_TipMon.Caption = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!CARCOF_TIPMON))
   
   l_int_Cof_TipMon = g_rst_Princi!CARCOF_TIPMON
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatInm()
   Dim r_str_TipVia  As String
   Dim r_str_TipZon  As String
   Dim r_str_Depart  As String
   Dim r_str_Provin  As String
   Dim r_str_Distri  As String
   
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SOLINM_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA))
   r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON))

   pnl_Inm_Direcc.Caption = r_str_TipVia & " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMERO)
   
   If Len(Trim(Trim(g_rst_Princi!SOLINM_INTDPT))) > 0 Then
      pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")"
   End If
   
   If Len(Trim(Trim(g_rst_Princi!SOLINM_NOMZON))) > 0 Then
      pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!SOLINM_NOMZON) & Chr(13) & Chr(10)
   Else
      pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & Chr(13) & Chr(10)
   End If
   
   'Departamento
   r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000")
   
   'Provincia
   r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00")
   
   'Distrito
   r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
   
   pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart

   
   pnl_Inm_TipPro.Caption = moddat_gf_Consulta_ParDes("221", CStr(g_rst_Princi!SOLINM_TIPPER))
   
   If g_rst_Princi!SOLINM_TIPPER = 2 Then
      'Persona Jurídica
      
      pnl_Inm_JurEmp.Caption = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_PROTDO)) & "-" & Trim(g_rst_Princi!SOLINM_PRONDO) & " / " & Trim(g_rst_Princi!SOLINM_PRORZS)
      pnl_Inm_JurRep.Caption = Trim(g_rst_Princi!SOLINM_PROAPP) & " " & Trim(g_rst_Princi!SOLINM_PROAPM) & " " & Trim(g_rst_Princi!SOLINM_PRONOM)
      
      r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_PROTVI))
      r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_PROTZO))
   
      pnl_Inm_JurDir.Caption = r_str_TipVia & " " & Trim(g_rst_Princi!SOLINM_PRONVI) & " " & Trim(g_rst_Princi!SOLINM_PRONUM)
      
      If Len(Trim(Trim(g_rst_Princi!SOLINM_PROINT))) > 0 Then
         pnl_Inm_JurDir.Caption = pnl_Inm_JurDir.Caption & " (" & Trim(g_rst_Princi!SOLINM_PROINT) & ")"
      End If
      
      If Len(Trim(Trim(g_rst_Princi!SOLINM_PRONZO))) > 0 Then
         pnl_Inm_JurDir.Caption = pnl_Inm_JurDir.Caption & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!SOLINM_PRONZO) & Chr(13) & Chr(10)
      Else
         pnl_Inm_JurDir.Caption = pnl_Inm_JurDir.Caption & Chr(13) & Chr(10)
      End If
      
      r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_PROUBI, 2) & "0000")
      r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_PROUBI, 4) & "00")
      r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_PROUBI))
      
      pnl_Inm_JurDir.Caption = pnl_Inm_JurDir.Caption & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   Else
      'Persona Natural
      
      pnl_Inm_NatTit.Caption = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_PROTDO)) & "-" & Trim(g_rst_Princi!SOLINM_PRONDO) & " / " & Trim(g_rst_Princi!SOLINM_PROAPP) & " " & Trim(g_rst_Princi!SOLINM_PROAPM) & " " & Trim(g_rst_Princi!SOLINM_PRONOM)
      
      If g_rst_Princi!SOLINM_CYGTDO > 0 Then
         pnl_Inm_NatCyg.Caption = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_CYGTDO)) & "-" & Trim(g_rst_Princi!SOLINM_CYGNDO) & " / " & Trim(g_rst_Princi!SOLINM_CYGAPP) & " " & Trim(g_rst_Princi!SOLINM_CYGAPM) & " " & Trim(g_rst_Princi!SOLINM_CYGNOM)
      End If
   End If


   l_str_Inm_CodPry = Trim(g_rst_Princi!SOLINM_PRYCOD & "")
   l_int_Inm_PryMCs = g_rst_Princi!SOLINM_PRYMCS
   l_int_Inm_TipVia = g_rst_Princi!SOLINM_TIPVIA
   l_str_Inm_NomVia = Trim(g_rst_Princi!SOLINM_NOMVIA & "")
   l_str_Inm_NumVia = Trim(g_rst_Princi!SOLINM_NUMERO & "")
   l_str_Inm_IntDpt = Trim(g_rst_Princi!SOLINM_INTDPT & "")
   l_int_Inm_TipZon = g_rst_Princi!SOLINM_TIPZON
   l_str_Inm_NomZon = Trim(g_rst_Princi!SOLINM_NOMZON & "")
   l_str_Inm_Refere = Trim(g_rst_Princi!SOLINM_REFERE & "")
   l_str_Inm_UbiGeo = Trim(g_rst_Princi!SOLINM_UBIGEO & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatAut()
   g_str_Parame = "SELECT * FROM TRA_AUTDES WHERE "
   g_str_Parame = g_str_Parame & "AUTDES_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   pnl_Aut_FueFin.Caption = moddat_gf_Consulta_ParDes("502", g_rst_Princi!AUTDES_FUEFIN)
   pnl_Aut_BonoBP.Caption = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!AUTDES_BONOBP))
   pnl_Aut_FecDes.Caption = gf_FormatoFecha(CStr(g_rst_Princi!AUTDES_FECDES))
   txt_Aut_Observ.Text = Trim(g_rst_Princi!AUTDES_OBSERV & "")
      
   l_int_Aut_BonoBP = g_rst_Princi!AUTDES_BONOBP
   l_str_Aut_FecDes = gf_FormatoFecha(CStr(g_rst_Princi!AUTDES_FECDES))
   l_str_Aut_FueFin = Trim(g_rst_Princi!AUTDES_FUEFIN & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_GenCro()
   Dim r_int_Contad     As Integer
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_SegPre     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_OtrCar     As Double
   Dim r_dbl_TotCuo     As Double
   Dim r_dbl_Comisi     As Double
   
   Call moddat_gs_FecSis
   
   'Calculando Seguro de Préstamo
   If l_int_Seg_AplPre = 1 Then           'Factor
      l_dbl_SegPre = CDbl(Format(l_dbl_Seg_FoiPre * l_dbl_Cre_MtoPre, "###,###,##0.00"))
   Else                                   'Importe
      l_dbl_SegPre = CDbl(Format(l_dbl_Seg_FoiPre, "###,###,##0.00"))
   End If
   
   'Calculando Seguro de Vivienda
   If l_int_Seg_AplViv = 1 Then           'Factor
      l_dbl_SegViv = CDbl(Format(l_dbl_Seg_FoiViv * l_dbl_Tas_ValFab, "###,###,##0.00"))
   Else                                   'Importe
      l_dbl_SegViv = CDbl(Format(l_dbl_Seg_FoiViv, "###,###,##0.00"))
   End If
   
   'Obteniendo Valor de Otros Cargos
   Call moddat_gf_Consulta_ParPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, "104", Format(moddat_g_int_TipMon, "000"))
   l_dbl_OtrCar = moddat_g_arr_Genera(1).Genera_Cantid
   
   
   'Guardando Monto Original de Préstamo en Moneda de Préstamo
   l_dbl_Cre_PreMPr = l_dbl_Cre_MtoPre
   
   'Inicializando Porcentajes de Tramo
   l_dbl_PorNCo = 100
   l_dbl_PorCon = 0
   
   'Inicializando Importes en Tramos
   l_dbl_ImpNCo = l_dbl_Cre_MtoPre
   l_dbl_ImpCon = 0
      
   'Si Producto es Mivivienda y tiene Bono Buen Pagador
   If moddat_g_str_CodPrd = "001" And l_int_Aut_BonoBP = 1 Then
      'Obteniendo Porcentaje Tramo Concesional
      Call moddat_gf_Consulta_ParPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, "701", "502")
      l_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
       
      'Obteniendo Porcentaje Tramo No Concesional
      Call moddat_gf_Consulta_ParPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, "701", "503")
      l_dbl_PorNCo = moddat_g_arr_Genera(1).Genera_Cantid
   
      'Calculando Importes por Tramos
      l_dbl_ImpNCo = CDbl(Format((l_dbl_PorNCo / 100) * l_dbl_Cre_MtoPre, "###,###,##0.00"))
      l_dbl_ImpCon = CDbl(Format((l_dbl_PorCon / 100) * l_dbl_Cre_MtoPre, "###,###,##0.00"))
      
      'Cambiando el Monto del Préstamo por el del Tramo No Concesional
      l_dbl_Cre_MtoPre = l_dbl_ImpNCo
   End If
   
   
   'Generando Cronograma Cliente No Concesional
   Call gs_Calcul_Cliente(l_arr_CliNCo(), l_dbl_Cre_TasInt, l_int_Cre_CuoExt, l_str_Aut_FecDes, l_int_Cre_NumCuo, l_dbl_ImpNCo, 1, l_int_Cre_PerGra, "")
   
   
   'Si Producto es MiVivienda se generan Cronogramas de COFIDE
   If moddat_g_str_CodPrd = "001" Then
      'Generando COFIDE Cronograma No Concesional
      Call gs_Calcul_COFIDE(l_arr_CofNCo(), l_dbl_Cre_TasCof, l_dbl_Cre_ComCof, l_int_Cre_CuoExt, l_str_Aut_FecDes, l_int_Cre_NumCuo, l_dbl_ImpNCo, 1, l_int_Cre_PerGra, "")
      
      If l_int_Aut_BonoBP = 1 Then
         'Generando Cronograma Cliente Concesional
         Call gs_Calcul_Cliente(l_arr_CliCon(), l_dbl_Cre_TasInt, l_int_Cre_CuoExt, l_str_Aut_FecDes, l_int_Cre_NumCuo / 6, l_dbl_ImpCon, 2, l_int_Cre_PerGra, l_arr_CliNCo(l_int_Cre_NumCuo).CuoCli_FecVct)
         
         'Generando COFIDE Cronograma Concesional
         Call gs_Calcul_COFIDE(l_arr_CofCon(), l_dbl_Cre_TasCof, l_dbl_Cre_ComCof, l_int_Cre_CuoExt, l_str_Aut_FecDes, l_int_Cre_NumCuo / 6, l_dbl_ImpCon, 2, l_int_Cre_PerGra, l_arr_CofNCo(l_int_Cre_NumCuo).CuoCof_FecVct)
      End If
   End If


   'Cliente No Concesional
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_SegPre = 0
   r_dbl_SegViv = 0
   r_dbl_OtrCar = 0
   r_dbl_TotCuo = 0
   
   grd_CliNCo_Listad.Redraw = False
   For r_int_Contad = 1 To UBound(l_arr_CliNCo)
      grd_CliNCo_Listad.Rows = grd_CliNCo_Listad.Rows + 1
      grd_CliNCo_Listad.Row = grd_CliNCo_Listad.Rows - 1
      
      'Número de Cuota
      grd_CliNCo_Listad.Col = 0
      grd_CliNCo_Listad.Text = Format(r_int_Contad, "000")
   
      'Fecha de Vencimiento
      grd_CliNCo_Listad.Col = 1
      grd_CliNCo_Listad.Text = l_arr_CliNCo(r_int_Contad).CuoCli_FecVct
   
      'Capital
      grd_CliNCo_Listad.Col = 2
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
      r_dbl_Capita = r_dbl_Capita + CDbl(grd_CliNCo_Listad)
      
      'Interes
      grd_CliNCo_Listad.Col = 3
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
      r_dbl_Intere = r_dbl_Intere + CDbl(grd_CliNCo_Listad)
   
      'Seguro Desgravamen
      grd_CliNCo_Listad.Col = 4
      grd_CliNCo_Listad.Text = Format(l_dbl_SegPre, "###,###,##0.00")
      r_dbl_SegPre = r_dbl_SegPre + CDbl(grd_CliNCo_Listad)
   
      'Seguro Vivienda
      grd_CliNCo_Listad.Col = 5
      grd_CliNCo_Listad.Text = Format(l_dbl_SegViv, "###,###,##0.00")
      r_dbl_SegViv = r_dbl_SegViv + CDbl(grd_CliNCo_Listad)
   
      'Otros Cargos
      grd_CliNCo_Listad.Col = 6
      grd_CliNCo_Listad.Text = Format(l_dbl_OtrCar, "###,###,##0.00")
      r_dbl_OtrCar = r_dbl_OtrCar + CDbl(grd_CliNCo_Listad)
   
      'Valor Cuota
      grd_CliNCo_Listad.Col = 7
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_Capita + l_arr_CliNCo(r_int_Contad).CuoCli_Intere + l_dbl_SegPre + l_dbl_SegViv + l_dbl_OtrCar, "###,###,##0.00")
      r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CliNCo_Listad)
   
      'Saldo Capital
      grd_CliNCo_Listad.Col = 8
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
   Next r_int_Contad
   grd_CliNCo_Listad.Redraw = True
   
   pnl_CliNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CliNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CliNCo_SegPre.Caption = Format(r_dbl_SegPre, "###,###,##0.00") & " "
   pnl_CliNCo_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
   pnl_CliNCo_OtrCar.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
   pnl_CliNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
   
   If grd_CliNCo_Listad.Rows > 0 Then
      Call gs_UbiIniGrid(grd_CliNCo_Listad)
   End If

   
   If moddat_g_str_CodPrd = "001" Then
      'Cofide No Concesional
      r_dbl_Capita = 0
      r_dbl_Intere = 0
      r_dbl_SegPre = 0
      r_dbl_SegViv = 0
      r_dbl_OtrCar = 0
      r_dbl_TotCuo = 0
      r_dbl_Comisi = 0
      
      grd_CofNCo_Listad.Redraw = False
      For r_int_Contad = 1 To UBound(l_arr_CofNCo)
         grd_CofNCo_Listad.Rows = grd_CofNCo_Listad.Rows + 1
         grd_CofNCo_Listad.Row = grd_CofNCo_Listad.Rows - 1
         
         'Número de Cuota
         grd_CofNCo_Listad.Col = 0
         grd_CofNCo_Listad.Text = Format(r_int_Contad, "000")
      
         'Fecha de Vencimiento
         grd_CofNCo_Listad.Col = 1
         grd_CofNCo_Listad.Text = l_arr_CofNCo(r_int_Contad).CuoCof_FecVct
      
         'Capital
         grd_CofNCo_Listad.Col = 2
         grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCof_Capita, "###,###,##0.00")
         r_dbl_Capita = r_dbl_Capita + CDbl(grd_CofNCo_Listad)
         
         'Interes
         grd_CofNCo_Listad.Col = 3
         grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCof_Intere, "###,###,##0.00")
         r_dbl_Intere = r_dbl_Intere + CDbl(grd_CofNCo_Listad)
      
         'Comisión
         grd_CofNCo_Listad.Col = 4
         grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCof_Comisi, "###,###,##0.00")
         r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_CofNCo_Listad)
      
         'Valor Cuota
         grd_CofNCo_Listad.Col = 5
         grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCof_Capita + l_arr_CofNCo(r_int_Contad).CuoCof_Intere + l_arr_CofNCo(r_int_Contad).CuoCof_Comisi, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CofNCo_Listad)
      
         'Saldo Capital
         grd_CofNCo_Listad.Col = 6
         grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCof_SalCap, "###,###,##0.00")
      Next r_int_Contad
      grd_CofNCo_Listad.Redraw = True
   
      pnl_CofNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
      pnl_CofNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
      pnl_CofNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
      pnl_CofNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
      
      If grd_CofNCo_Listad.Rows > 0 Then
         Call gs_UbiIniGrid(grd_CofNCo_Listad)
      End If
   
      If l_int_Aut_BonoBP = 1 Then
         'Cliente Concesional
         r_dbl_Capita = 0
         r_dbl_Intere = 0
         r_dbl_SegPre = 0
         r_dbl_SegViv = 0
         r_dbl_OtrCar = 0
         r_dbl_TotCuo = 0
         
         grd_CliCon_Listad.Redraw = False
         For r_int_Contad = 1 To UBound(l_arr_CliCon)
            grd_CliCon_Listad.Rows = grd_CliCon_Listad.Rows + 1
            grd_CliCon_Listad.Row = grd_CliCon_Listad.Rows - 1
            
            'Número de Cuota
            grd_CliCon_Listad.Col = 0
            grd_CliCon_Listad.Text = Format(r_int_Contad, "000")
         
            'Fecha de Vencimiento
            grd_CliCon_Listad.Col = 1
            grd_CliCon_Listad.Text = l_arr_CliCon(r_int_Contad).CuoCli_FecVct
         
            'Capital
            grd_CliCon_Listad.Col = 2
            grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
            r_dbl_Capita = r_dbl_Capita + CDbl(grd_CliCon_Listad)
            
            'Interes
            grd_CliCon_Listad.Col = 3
            grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
            r_dbl_Intere = r_dbl_Intere + CDbl(grd_CliCon_Listad)
         
            'Valor Cuota
            grd_CliCon_Listad.Col = 4
            grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_Capita + l_arr_CliCon(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CliCon_Listad)
         
            'Saldo Capital
            grd_CliCon_Listad.Col = 5
            grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
         Next r_int_Contad
         grd_CliCon_Listad.Redraw = True
         
         pnl_CliCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
         pnl_CliCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
         pnl_CliCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
         
         If grd_CliCon_Listad.Rows > 0 Then
            Call gs_UbiIniGrid(grd_CliCon_Listad)
         End If
   
   
         'Cofide Concesional
         r_dbl_Capita = 0
         r_dbl_Intere = 0
         r_dbl_SegPre = 0
         r_dbl_SegViv = 0
         r_dbl_OtrCar = 0
         r_dbl_TotCuo = 0
         r_dbl_Comisi = 0
         
         grd_CofCon_Listad.Redraw = False
         For r_int_Contad = 1 To UBound(l_arr_CofCon)
            grd_CofCon_Listad.Rows = grd_CofCon_Listad.Rows + 1
            grd_CofCon_Listad.Row = grd_CofCon_Listad.Rows - 1
            
            'Número de Cuota
            grd_CofCon_Listad.Col = 0
            grd_CofCon_Listad.Text = Format(r_int_Contad, "000")
         
            'Fecha de Vencimiento
            grd_CofCon_Listad.Col = 1
            grd_CofCon_Listad.Text = l_arr_CofCon(r_int_Contad).CuoCof_FecVct
         
            'Capital
            grd_CofCon_Listad.Col = 2
            grd_CofCon_Listad.Text = Format(l_arr_CofCon(r_int_Contad).CuoCof_Capita, "###,###,##0.00")
            r_dbl_Capita = r_dbl_Capita + CDbl(grd_CofCon_Listad)
            
            'Interes
            grd_CofCon_Listad.Col = 3
            grd_CofCon_Listad.Text = Format(l_arr_CofCon(r_int_Contad).CuoCof_Intere, "###,###,##0.00")
            r_dbl_Intere = r_dbl_Intere + CDbl(grd_CofCon_Listad)
         
            'Comisión
            grd_CofCon_Listad.Col = 4
            grd_CofCon_Listad.Text = Format(l_arr_CofCon(r_int_Contad).CuoCof_Comisi, "###,###,##0.00")
            r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_CofCon_Listad)
         
            'Valor Cuota
            grd_CofCon_Listad.Col = 5
            grd_CofCon_Listad.Text = Format(l_arr_CofCon(r_int_Contad).CuoCof_Capita + l_arr_CofCon(r_int_Contad).CuoCof_Intere + l_arr_CofCon(r_int_Contad).CuoCof_Comisi, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CofCon_Listad)
         
            'Saldo Capital
            grd_CofCon_Listad.Col = 6
            grd_CofCon_Listad.Text = Format(l_arr_CofCon(r_int_Contad).CuoCof_SalCap, "###,###,##0.00")
         Next r_int_Contad
         grd_CofCon_Listad.Redraw = True
         
         pnl_CofCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
         pnl_CofCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
         pnl_CofCon_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
         pnl_CofCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
         
         If grd_CofCon_Listad.Rows > 0 Then
            Call gs_UbiIniGrid(grd_CofCon_Listad)
         End If
      End If
   End If
End Sub

Private Sub fs_Genera_Operac()
   'Obteniendo Datos de Maestro de Productos
   Call fs_Buscar_MaePrd
   Call fs_Buscar_DatCro
      
   'Actualizando Tablas
   Call fs_Graba_ActDir
   Call fs_Graba_Seguim
   
   'Grabando Creditos
   Call fs_Graba_Credit
   Call fs_Graba_Cronog
End Sub

Private Sub fs_Buscar_MaePrd()
   g_str_Parame = "SELECT * FROM CRE_PRODUC WHERE "
   g_str_Parame = g_str_Parame & "PRODUC_CODIGO = '" & moddat_g_str_CodPrd & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
   
      l_int_ClaCre = g_rst_Princi!PRODUC_CODCLA
      l_int_IndITF = g_rst_Princi!PRODUC_INDITF
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatCro()
   '1ra Cuota Cronograma Cliente No Concesional
   grd_CliNCo_Listad.Row = 0
   grd_CliNCo_Listad.Col = 1:    l_str_CliNCo_PriVct = grd_CliNCo_Listad.Text
   grd_CliNCo_Listad.Col = 7:    l_dbl_CliNCo_CuoIni = CDbl(grd_CliNCo_Listad.Text)
   
   'Ultima Cuota Cronograma Cliente No Concesional
   grd_CliNCo_Listad.Row = grd_CliNCo_Listad.Rows - 1
   grd_CliNCo_Listad.Col = 1:    l_str_CliNCo_UltVct = grd_CliNCo_Listad.Text
   grd_CliNCo_Listad.Col = 7:    l_dbl_CliNCo_CuoFin = CDbl(grd_CliNCo_Listad.Text)
   
   'Cuota Fija Cronograma Cliente No Concesional
   grd_CliNCo_Listad.Row = 1
   grd_CliNCo_Listad.Col = 1:    l_str_CliNCo_PrxVct = grd_CliNCo_Listad.Text
   grd_CliNCo_Listad.Col = 7:    l_dbl_CliNCo_CuoFij = CDbl(grd_CliNCo_Listad.Text)
   
   'Si Producto es Mivivienda
   If moddat_g_str_CodPrd = "001" Then
      '1ra Cuota Cronograma Cofide No Concesional
      grd_CofNCo_Listad.Row = 0
      grd_CofNCo_Listad.Col = 5:    l_dbl_CofNCo_CuoIni = CDbl(grd_CofNCo_Listad.Text)
   
      'Cuota Fija Cronograma Cofide No Concesional
      grd_CofNCo_Listad.Row = 1
      grd_CofNCo_Listad.Col = 5:    l_dbl_CofNCo_CuoFij = CDbl(grd_CofNCo_Listad.Text)
   
      'Cuota Final Cronograma Cofide No Concesional
      grd_CofNCo_Listad.Row = grd_CofNCo_Listad.Rows - 1
      grd_CofNCo_Listad.Col = 5:    l_dbl_CofNCo_CuoFin = CDbl(grd_CofNCo_Listad.Text)
      
      If l_int_Aut_BonoBP = 1 Then
         'Cuota Inicial Cronograma Cofide No Concesional
         grd_CofCon_Listad.Row = 0
         grd_CofCon_Listad.Col = 5:    l_dbl_CofCon_CuoIni = CDbl(grd_CofCon_Listad.Text)
      
         'Cuota Fija Cronograma Cofide No Concesional
         grd_CofCon_Listad.Row = 1
         grd_CofCon_Listad.Col = 5:    l_dbl_CofCon_CuoFij = CDbl(grd_CofCon_Listad.Text)

         'Cuota Final Cronograma Cofide No Concesional
         grd_CofCon_Listad.Row = grd_CofCon_Listad.Rows - 1
         grd_CofCon_Listad.Col = 5:    l_dbl_CofCon_CuoFin = CDbl(grd_CofCon_Listad.Text)
      End If
   End If
End Sub

Private Sub fs_Graba_Credit()
   'Generando Número de Operación
   l_str_NumOpe = ff_Genera_NumOpe()
   
   'Grabando Cabecera de Credito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPMAE_CREA ("
      
      g_str_Parame = g_str_Parame & "'" & l_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodMod & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodEje & "', "
      g_str_Parame = g_str_Parame & l_str_Cli_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_Inm_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Inm_CodPry & "', "
      g_str_Parame = g_str_Parame & CStr(l_int_Inm_PryMCs) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_Aut_FueFin & "', "
      g_str_Parame = g_str_Parame & CStr(l_int_ClaCre) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_Cre_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_Cre_CuoExt) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_Cre_PerGra) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_Cre_NumCuo) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(l_str_Aut_FecDes), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(CDate(l_str_Leg_AprCom), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(CDate(l_str_CliNCo_PriVct), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(CDate(l_str_CliNCo_PrxVct), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(CDate(l_str_CliNCo_UltVct), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_MtoPre) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_PreMPr) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_PreSol) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_PreDol) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_ComVta) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_ApoPro) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CliNCo_CuoIni) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CliNCo_CuoFin) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CliNCo_CuoFij) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_MtoPre) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_TasInt) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_Seg_EmpDes & "', "
      g_str_Parame = g_str_Parame & CStr(l_int_Seg_TipSeg) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_Seg_EmpViv & "', "
      g_str_Parame = g_str_Parame & CStr(l_int_Seg_AplPre) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Seg_FoiPre) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_Seg_AplViv) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Seg_FoiViv) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_SegViv) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_SegPre) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_OtrCar) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_Tas_EmpPer & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Tas_NomPer & "', "
      g_str_Parame = g_str_Parame & CStr(l_int_Tas_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Tas_ValCom) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Tas_ValFab) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_TCaDol) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_TCaMPr) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_Cof_NumOpe & "', "
      g_str_Parame = g_str_Parame & CStr(l_int_Cof_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cof_MtoDes) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_Aut_BonoBP) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_TasCof) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Cre_ComCof) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CofNCo_CuoIni) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CofNCo_CuoFij) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CofNCo_CuoFin) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ImpNCo) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_PorNCo) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ImpNCo) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CofCon_CuoIni) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CofCon_CuoFij) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CofCon_CuoFin) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ImpCon) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_PorCon) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ImpCon) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_IndITF) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                           'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRESOLMAE. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
End Sub

Private Sub fs_Graba_Cronog()
   Dim r_int_Contad  As Integer
   Dim r_int_NumCuo  As Integer
   Dim r_str_FecVct  As String
   Dim r_dbl_Capita  As Double
   Dim r_dbl_Intere  As Double
   Dim r_dbl_Comisi  As Double
   Dim r_dbl_SegDes  As Double
   Dim r_dbl_SegViv  As Double
   Dim r_dbl_OtrCar  As Double
   Dim r_dbl_SalCap  As Double
   
   'Grabando Cronograma Cliente No Concesional
   grd_CliNCo_Listad.Redraw = False
   For r_int_Contad = 0 To grd_CliNCo_Listad.Rows - 1
      grd_CliNCo_Listad.Row = r_int_Contad
   
      grd_CliNCo_Listad.Col = 0:          r_int_NumCuo = CInt(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 1:          r_str_FecVct = grd_CliNCo_Listad.Text
      grd_CliNCo_Listad.Col = 2:          r_dbl_Capita = CDbl(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 3:          r_dbl_Intere = CDbl(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 4:          r_dbl_SegDes = CDbl(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 5:          r_dbl_SegViv = CDbl(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 6:          r_dbl_OtrCar = CDbl(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 8:          r_dbl_SalCap = CDbl(grd_CliNCo_Listad.Text)
      
      If Not ff_Inserta_HipCuo(l_str_NumOpe, 1, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, r_dbl_SegDes, r_dbl_SegViv, r_dbl_OtrCar, r_dbl_SalCap) Then
         Exit Sub
      End If
      
   Next r_int_Contad
   
   grd_CliNCo_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_CliNCo_Listad)

   'Si Producto es MiVivienda
   If moddat_g_str_CodPrd = "001" Then
      'Grabando Cronograma Cofide No Concesional
      For r_int_Contad = 0 To grd_CofNCo_Listad.Rows - 1
         grd_CofNCo_Listad.Row = r_int_Contad
      
         grd_CofNCo_Listad.Col = 0:       r_int_NumCuo = CInt(grd_CofNCo_Listad.Text)
         grd_CofNCo_Listad.Col = 1:       r_str_FecVct = grd_CofNCo_Listad.Text
         grd_CofNCo_Listad.Col = 2:       r_dbl_Capita = CDbl(grd_CofNCo_Listad.Text)
         grd_CofNCo_Listad.Col = 3:       r_dbl_Intere = CDbl(grd_CofNCo_Listad.Text)
         grd_CofNCo_Listad.Col = 4:       r_dbl_Comisi = CDbl(grd_CofNCo_Listad.Text)
         grd_CofNCo_Listad.Col = 6:       r_dbl_SalCap = CDbl(grd_CofNCo_Listad.Text)
         
         If Not ff_Inserta_HipCuo(l_str_NumOpe, 3, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, r_dbl_Comisi, 0, 0, 0, r_dbl_SalCap) Then
            Exit Sub
         End If
      Next r_int_Contad
      
      'Si tiene Premio Bono Buen Pagador
      If l_int_Aut_BonoBP = 1 Then
         'Grabando Cronograma Cliente Concesional
         grd_CliCon_Listad.Redraw = False
         For r_int_Contad = 0 To grd_CliCon_Listad.Rows - 1
            grd_CliCon_Listad.Row = r_int_Contad
         
            grd_CliCon_Listad.Col = 0:       r_int_NumCuo = CInt(grd_CliCon_Listad.Text)
            grd_CliCon_Listad.Col = 1:       r_str_FecVct = grd_CliCon_Listad.Text
            grd_CliCon_Listad.Col = 2:       r_dbl_Capita = CDbl(grd_CliCon_Listad.Text)
            grd_CliCon_Listad.Col = 3:       r_dbl_Intere = CDbl(grd_CliCon_Listad.Text)
            grd_CliCon_Listad.Col = 5:       r_dbl_SalCap = CDbl(grd_CliCon_Listad.Text)
            
            If Not ff_Inserta_HipCuo(l_str_NumOpe, 2, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, 0, r_dbl_SalCap) Then
               Exit Sub
            End If
         Next r_int_Contad
         grd_CliCon_Listad.Redraw = True
         Call gs_UbiIniGrid(grd_CliCon_Listad)
         
         'Grabando Cronograma Cofide Concesional
         For r_int_Contad = 0 To grd_CofCon_Listad.Rows - 1
            grd_CofCon_Listad.Row = r_int_Contad
         
            grd_CofCon_Listad.Col = 0:       r_int_NumCuo = CInt(grd_CofCon_Listad.Text)
            grd_CofCon_Listad.Col = 1:       r_str_FecVct = grd_CofCon_Listad.Text
            grd_CofCon_Listad.Col = 2:       r_dbl_Capita = CDbl(grd_CofCon_Listad.Text)
            grd_CofCon_Listad.Col = 3:       r_dbl_Intere = CDbl(grd_CofCon_Listad.Text)
            grd_CofCon_Listad.Col = 4:       r_dbl_Comisi = CDbl(grd_CofCon_Listad.Text)
            grd_CofCon_Listad.Col = 6:       r_dbl_SalCap = CDbl(grd_CofCon_Listad.Text)
            
            If Not ff_Inserta_HipCuo(l_str_NumOpe, 4, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, r_dbl_Comisi, 0, 0, 0, r_dbl_SalCap) Then
               Exit Sub
            End If
         Next r_int_Contad
      End If
   End If
End Sub

Private Sub fs_Graba_ActDir()
   'Actualizando Dirección en Solicitud de Crédito (Se guarda la Dirección Actual)
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_SOLMAE_DESEMB ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & CStr(l_int_Cli_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_Cli_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Cli_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Cli_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(l_int_Cli_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_Cli_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Cli_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Cli_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_Cli_Telefo & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                           'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_SOLMAE_DESEMB. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop


   'Actualizando Dirección en Maestro de Clientes (Titular) (Se guarda la Dirección del Inmueble)
   'moddat_g_int_FlgGOK = False
   'moddat_g_int_CntErr = 0
   
   'Do While moddat_g_int_FlgGOK = False
   '   g_str_Parame = "USP_CLI_DATGEN_DIRECC ("
   '
   '   g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
   '   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
   '   g_str_Parame = g_str_Parame & CStr(l_int_Inm_TipVia) & ", "
   '   g_str_Parame = g_str_Parame & "'" & l_str_Inm_NomVia & "', "
   '   g_str_Parame = g_str_Parame & "'" & l_str_Inm_NumVia & "', "
   '   g_str_Parame = g_str_Parame & "'" & l_str_Inm_IntDpt & "', "
   '   g_str_Parame = g_str_Parame & CStr(l_int_Inm_TipZon) & ", "
   '   g_str_Parame = g_str_Parame & "'" & l_str_Inm_NomZon & "', "
   '   g_str_Parame = g_str_Parame & "'" & l_str_Inm_Refere & "', "
   '   g_str_Parame = g_str_Parame & "'" & l_str_Inm_UbiGeo & "', "
   '
   '   'Datos de Auditoria
   '   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
   '   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
   '   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
   '   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                           'Código Sucursal
   '   g_str_Parame = g_str_Parame & "1)"
   '
   '   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
   '      moddat_g_int_CntErr = moddat_g_int_CntErr + 1
   '   Else
   '      moddat_g_int_FlgGOK = True
   '   End If

   '   If moddat_g_int_CntErr = 6 Then
   '      If MsgBox("No se pudo completar el procedimiento USP_CLI_DATGEN_DIRECC. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
   '         Exit Sub
   '      Else
   '         moddat_g_int_CntErr = 0
   '      End If
   '   End If
   'Loop


   'Actualizando Dirección en Maestro de Clientes (Cónyuge) (Se guarda la Dirección del Inmueble)
   'If moddat_g_int_CygTDo > 0 Then
   '   moddat_g_int_FlgGOK = False
   '   moddat_g_int_CntErr = 0
   '
   '   Do While moddat_g_int_FlgGOK = False
   '      g_str_Parame = "USP_CLI_DATGEN_DIRECC ("
   '
   '      g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
   '      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
   '      g_str_Parame = g_str_Parame & CStr(l_int_Inm_TipVia) & ", "
   '      g_str_Parame = g_str_Parame & "'" & l_str_Inm_NomVia & "', "
   '      g_str_Parame = g_str_Parame & "'" & l_str_Inm_NumVia & "', "
   '      g_str_Parame = g_str_Parame & "'" & l_str_Inm_IntDpt & "', "
   '      g_str_Parame = g_str_Parame & CStr(l_int_Inm_TipZon) & ", "
   '      g_str_Parame = g_str_Parame & "'" & l_str_Inm_NomZon & "', "
   '      g_str_Parame = g_str_Parame & "'" & l_str_Inm_Refere & "', "
   '      g_str_Parame = g_str_Parame & "'" & l_str_Inm_UbiGeo & "', "
         
         'Datos de Auditoria
   '      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
   '      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
   '      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
   '      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                           'Código Sucursal
   '      g_str_Parame = g_str_Parame & "1)"
         
   '      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
   '         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
   '      Else
   '         moddat_g_int_FlgGOK = True
   '      End If
   
   '      If moddat_g_int_CntErr = 6 Then
   '         If MsgBox("No se pudo completar el procedimiento USP_CLI_DATGEN_DIRECC. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
   '            Exit Sub
   '         Else
   '            moddat_g_int_CntErr = 0
   '         End If
   '      End If
   '   Loop
   'End If
End Sub

Private Sub fs_Graba_Seguim()
   Dim r_int_DiaTra     As Integer
   
   'Obteniendo Fecha de Inicio de Instancia
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(modatecli_g_con_Desemb)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(gf_FormatoFecha(g_rst_Princi!SEGUIM_FECINI)))
   Else
      r_int_DiaTra = 0
   End If
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_Desemb, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_Desemb, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
End Sub

Private Function ff_Inserta_HipCuo(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_NumCuo As Integer, ByVal p_FecVct As String, ByVal p_Capita As Double, ByVal p_Intere As Double, ByVal p_Comisi As Double, ByVal p_SegDes As Double, ByVal p_SegViv As Double, ByVal p_OtrGas As Double, ByVal p_SalCap As Double) As Integer
   ff_Inserta_HipCuo = False
   
   'Grabando Cabecera de Credito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPCUO_CREA ("
      
      g_str_Parame = g_str_Parame & "'" & p_NumOpe & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipCro) & ", "
      g_str_Parame = g_str_Parame & CStr(p_NumCuo) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(p_FecVct), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(p_Capita) & ", "
      g_str_Parame = g_str_Parame & CStr(p_Intere) & ", "
      g_str_Parame = g_str_Parame & CStr(p_Comisi) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SegDes) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SegViv) & ", "
      g_str_Parame = g_str_Parame & CStr(p_OtrGas) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SalCap) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                           'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop

   ff_Inserta_HipCuo = True
End Function

Private Function ff_Genera_NumOpe() As String
   Dim r_lng_NumSol     As Long
   Dim r_str_NumSol     As String
   
   ff_Genera_NumOpe = ""
   
   'Obteniendo Número de Solicitud
   Call moddat_gs_FecSis
   
   g_str_Parame = "SELECT * FROM CRE_FOLIOS WHERE "
   g_str_Parame = g_str_Parame & "FOLIOS_TIPFOL = 2 AND "
   g_str_Parame = g_str_Parame & "FOLIOS_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "FOLIOS_CODSUC = '000' AND "
   g_str_Parame = g_str_Parame & "FOLIOS_PERANO = " & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      r_lng_NumSol = 1
   Else
      r_lng_NumSol = g_rst_Genera!FOLIOS_NUMERO + 1
   End If

   r_str_NumSol = moddat_g_str_CodPrd & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & Format(r_lng_NumSol, "00000")
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      'Actualizando Correlativo
      g_str_Parame = "USP_CRE_FOLIOS ("
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'000', "
      g_str_Parame = g_str_Parame & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & ", "
      g_str_Parame = g_str_Parame & CStr(r_lng_NumSol) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      
      g_str_Parame = g_str_Parame & "1, "
      
      If r_lng_NumSol = 1 Then
         g_str_Parame = g_str_Parame & "1) "
      Else
         g_str_Parame = g_str_Parame & "2) "
      End If
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_FOLIOS. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Genera_NumOpe = r_str_NumSol
End Function

