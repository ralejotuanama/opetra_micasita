VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Sim_OpeHip_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8970
   ClientLeft      =   825
   ClientTop       =   1245
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_288.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   15849
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
         Height          =   2865
         Left            =   30
         TabIndex        =   54
         Top             =   6060
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5054
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
            Height          =   2745
            Left            =   60
            TabIndex        =   55
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   4842
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Cliente (TNC)"
            TabPicture(0)   =   "OpeTra_frm_288.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "pnl_CliNCo_TotCuo"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "pnl_CliNCo_OtrCar"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel62"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel61"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel59"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel36"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel35"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel34"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel33"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel3"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "grd_CliNCo_Listad"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "pnl_CliNCo_Intere"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "pnl_CliNCo_SegPre"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_CliNCo_SegViv"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_CliNCo_Capita"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "SSPanel30"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).ControlCount=   16
            TabCaption(1)   =   "Cliente (TC)"
            TabPicture(1)   =   "OpeTra_frm_288.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "pnl_CliCon_TotCuo"
            Tab(1).Control(1)=   "SSPanel5"
            Tab(1).Control(2)=   "grd_CliCon_Listad"
            Tab(1).Control(3)=   "SSPanel8"
            Tab(1).Control(4)=   "SSPanel11"
            Tab(1).Control(5)=   "SSPanel12"
            Tab(1).Control(6)=   "SSPanel13"
            Tab(1).Control(7)=   "SSPanel21"
            Tab(1).Control(8)=   "pnl_CliCon_Intere"
            Tab(1).Control(9)=   "pnl_CliCon_Capita"
            Tab(1).ControlCount=   10
            TabCaption(2)   =   "MVI-Cofide (TNC)"
            TabPicture(2)   =   "OpeTra_frm_288.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_MViNCo_Listad"
            Tab(2).Control(1)=   "SSPanel14"
            Tab(2).Control(2)=   "SSPanel38"
            Tab(2).Control(3)=   "SSPanel41"
            Tab(2).Control(4)=   "SSPanel42"
            Tab(2).Control(5)=   "SSPanel43"
            Tab(2).Control(6)=   "SSPanel44"
            Tab(2).Control(7)=   "SSPanel45"
            Tab(2).Control(8)=   "SSPanel46"
            Tab(2).Control(9)=   "SSPanel47"
            Tab(2).Control(10)=   "SSPanel49"
            Tab(2).Control(11)=   "pnl_MViNCo_Capita"
            Tab(2).Control(12)=   "pnl_MViNCo_SegViv"
            Tab(2).Control(13)=   "pnl_MViNCo_SegPre"
            Tab(2).Control(14)=   "pnl_MViNCo_Intere"
            Tab(2).Control(15)=   "pnl_MViNCo_OtrCar"
            Tab(2).Control(16)=   "pnl_MViNCo_TotCuo"
            Tab(2).Control(17)=   "pnl_MViNCo_Comisi"
            Tab(2).ControlCount=   18
            TabCaption(3)   =   "MVI-Cofide (TC)"
            TabPicture(3)   =   "OpeTra_frm_288.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "pnl_MViCon_TotCuo"
            Tab(3).Control(1)=   "SSPanel15"
            Tab(3).Control(2)=   "grd_MviCon_Listad"
            Tab(3).Control(3)=   "SSPanel16"
            Tab(3).Control(4)=   "SSPanel17"
            Tab(3).Control(5)=   "SSPanel18"
            Tab(3).Control(6)=   "SSPanel19"
            Tab(3).Control(7)=   "SSPanel20"
            Tab(3).Control(8)=   "SSPanel22"
            Tab(3).Control(9)=   "pnl_MViCon_Intere"
            Tab(3).Control(10)=   "pnl_MViCon_Capita"
            Tab(3).Control(11)=   "pnl_MViCon_Comisi"
            Tab(3).ControlCount=   12
            TabCaption(4)   =   "Cofide"
            TabPicture(4)   =   "OpeTra_frm_288.frx":007C
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "pnl_CofNCo_TotCuo"
            Tab(4).Control(1)=   "SSPanel55"
            Tab(4).Control(2)=   "grd_CofNCo_Listad"
            Tab(4).Control(3)=   "SSPanel56"
            Tab(4).Control(4)=   "SSPanel58"
            Tab(4).Control(5)=   "SSPanel60"
            Tab(4).Control(6)=   "SSPanel63"
            Tab(4).Control(7)=   "SSPanel64"
            Tab(4).Control(8)=   "SSPanel65"
            Tab(4).Control(9)=   "pnl_CofNCo_Intere"
            Tab(4).Control(10)=   "pnl_CofNCo_Capita"
            Tab(4).Control(11)=   "pnl_CofNCo_Comisi"
            Tab(4).ControlCount=   12
            Begin Threed.SSPanel SSPanel30 
               Height          =   285
               Left            =   3450
               TabIndex        =   56
               Top             =   390
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel pnl_CliNCo_Capita 
               Height          =   285
               Left            =   2280
               TabIndex        =   57
               Top             =   2400
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
               Left            =   5790
               TabIndex        =   58
               Top             =   2400
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
               Left            =   4620
               TabIndex        =   59
               Top             =   2400
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
               Left            =   3450
               TabIndex        =   60
               Top             =   2400
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
               Height          =   1695
               Left            =   30
               TabIndex        =   61
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
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
            Begin Threed.SSPanel SSPanel3 
               Height          =   285
               Left            =   60
               TabIndex        =   62
               Top             =   390
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel33 
               Height          =   285
               Left            =   840
               TabIndex        =   63
               Top             =   390
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel34 
               Height          =   285
               Left            =   2280
               TabIndex        =   64
               Top             =   390
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel35 
               Height          =   285
               Left            =   8130
               TabIndex        =   65
               Top             =   390
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel36 
               Height          =   285
               Left            =   9420
               TabIndex        =   66
               Top             =   390
               Width           =   1560
               _Version        =   65536
               _ExtentX        =   2752
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel59 
               Height          =   285
               Left            =   4620
               TabIndex        =   67
               Top             =   390
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
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
            Begin Threed.SSPanel SSPanel61 
               Height          =   285
               Left            =   5790
               TabIndex        =   68
               Top             =   390
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
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
            Begin Threed.SSPanel SSPanel62 
               Height          =   285
               Left            =   6960
               TabIndex        =   69
               Top             =   390
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Portes"
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
            Begin Threed.SSPanel pnl_CliNCo_OtrCar 
               Height          =   285
               Left            =   6960
               TabIndex        =   70
               Top             =   2400
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
               Left            =   8130
               TabIndex        =   71
               Top             =   2400
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliCon_TotCuo 
               Height          =   285
               Left            =   -68370
               TabIndex        =   72
               Top             =   2400
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   285
               Left            =   -70530
               TabIndex        =   73
               Top             =   390
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interes"
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliCon_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   74
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
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
            Begin Threed.SSPanel SSPanel8 
               Height          =   285
               Left            =   -74940
               TabIndex        =   75
               Top             =   390
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
               Left            =   -74190
               TabIndex        =   76
               Top             =   390
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
               Left            =   -72690
               TabIndex        =   77
               Top             =   390
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel13 
               Height          =   285
               Left            =   -68370
               TabIndex        =   78
               Top             =   390
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -66210
               TabIndex        =   79
               Top             =   390
               Width           =   2235
               _Version        =   65536
               _ExtentX        =   3942
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel pnl_CliCon_Intere 
               Height          =   285
               Left            =   -70530
               TabIndex        =   80
               Top             =   2400
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
               Left            =   -72690
               TabIndex        =   81
               Top             =   2400
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin MSFlexGridLib.MSFlexGrid grd_MViNCo_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   82
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   10
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   -71790
               TabIndex        =   83
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel SSPanel38 
               Height          =   285
               Left            =   -74940
               TabIndex        =   84
               Top             =   390
               Width           =   705
               _Version        =   65536
               _ExtentX        =   1244
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel41 
               Height          =   285
               Left            =   -74250
               TabIndex        =   85
               Top             =   390
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel42 
               Height          =   285
               Left            =   -72840
               TabIndex        =   86
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel43 
               Height          =   285
               Left            =   -66390
               TabIndex        =   87
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel44 
               Height          =   285
               Left            =   -65310
               TabIndex        =   88
               Top             =   390
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel45 
               Height          =   285
               Left            =   -70710
               TabIndex        =   89
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
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
            Begin Threed.SSPanel SSPanel46 
               Height          =   285
               Left            =   -69630
               TabIndex        =   90
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
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
            Begin Threed.SSPanel SSPanel47 
               Height          =   285
               Left            =   -68550
               TabIndex        =   91
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Portes"
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
            Begin Threed.SSPanel SSPanel49 
               Height          =   285
               Left            =   -67470
               TabIndex        =   92
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "C. COFIDE"
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
            Begin Threed.SSPanel pnl_MViNCo_Capita 
               Height          =   285
               Left            =   -72840
               TabIndex        =   93
               Top             =   2400
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_SegViv 
               Height          =   285
               Left            =   -69630
               TabIndex        =   94
               Top             =   2400
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_SegPre 
               Height          =   285
               Left            =   -70710
               TabIndex        =   95
               Top             =   2400
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_Intere 
               Height          =   285
               Left            =   -71790
               TabIndex        =   96
               Top             =   2400
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_OtrCar 
               Height          =   285
               Left            =   -68550
               TabIndex        =   97
               Top             =   2400
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_TotCuo 
               Height          =   285
               Left            =   -66390
               TabIndex        =   98
               Top             =   2400
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_Comisi 
               Height          =   285
               Left            =   -67470
               TabIndex        =   99
               Top             =   2400
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViCon_TotCuo 
               Height          =   285
               Left            =   -67470
               TabIndex        =   100
               Top             =   2400
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -70950
               TabIndex        =   101
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin MSFlexGridLib.MSFlexGrid grd_MviCon_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   102
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
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
            Begin Threed.SSPanel SSPanel16 
               Height          =   285
               Left            =   -74940
               TabIndex        =   103
               Top             =   390
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   285
               Left            =   -74190
               TabIndex        =   104
               Top             =   390
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -72690
               TabIndex        =   105
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
               Left            =   -67470
               TabIndex        =   106
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
               Left            =   -65730
               TabIndex        =   107
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel22 
               Height          =   285
               Left            =   -69210
               TabIndex        =   108
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
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
            Begin Threed.SSPanel pnl_MViCon_Intere 
               Height          =   285
               Left            =   -70950
               TabIndex        =   109
               Top             =   2400
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViCon_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   110
               Top             =   2400
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViCon_Comisi 
               Height          =   285
               Left            =   -69210
               TabIndex        =   111
               Top             =   2400
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CofNCo_TotCuo 
               Height          =   285
               Left            =   -67500
               TabIndex        =   112
               Top             =   2400
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel55 
               Height          =   285
               Left            =   -70980
               TabIndex        =   113
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin MSFlexGridLib.MSFlexGrid grd_CofNCo_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   114
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
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
            Begin Threed.SSPanel SSPanel56 
               Height          =   285
               Left            =   -74940
               TabIndex        =   115
               Top             =   390
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel58 
               Height          =   285
               Left            =   -74220
               TabIndex        =   116
               Top             =   390
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel60 
               Height          =   285
               Left            =   -72720
               TabIndex        =   117
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel63 
               Height          =   285
               Left            =   -67500
               TabIndex        =   118
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel64 
               Height          =   285
               Left            =   -65760
               TabIndex        =   119
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel65 
               Height          =   285
               Left            =   -69240
               TabIndex        =   120
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
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
            Begin Threed.SSPanel pnl_CofNCo_Intere 
               Height          =   285
               Left            =   -70980
               TabIndex        =   121
               Top             =   2400
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
               Left            =   -72720
               TabIndex        =   122
               Top             =   2400
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CofNCo_Comisi 
               Height          =   285
               Left            =   -69240
               TabIndex        =   123
               Top             =   2400
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   765
         Left            =   30
         TabIndex        =   1
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
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   3915
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   9345
         End
         Begin EditLib.fpDoubleSingle ipp_TipCam 
            Height          =   315
            Left            =   10020
            TabIndex        =   43
            Top             =   390
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
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
         Begin VB.Label Label9 
            Caption         =   "Tipo de Cambio:"
            Height          =   315
            Left            =   8010
            TabIndex        =   42
            Top             =   390
            Width           =   1245
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Left            =   9450
            TabIndex        =   41
            Top             =   390
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   90
            TabIndex        =   5
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   4
            Top             =   60
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   6
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   615
            Left            =   630
            TabIndex        =   7
            Top             =   30
            Width           =   8505
            _Version        =   65536
            _ExtentX        =   15002
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "Simulación Operativa de Créditos Hipotecarios"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   10920
            Top             =   60
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_288.frx":0098
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   645
         Left            =   30
         TabIndex        =   8
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_288.frx":03A2
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_288.frx":07E4
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_288.frx":0AEE
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Imprimir Simulación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CalCuo 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_288.frx":0DF8
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Calcular Cuota"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   3765
         Left            =   30
         TabIndex        =   13
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6641
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
         Begin VB.ComboBox cmb_TipSeg 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1740
            Width           =   3915
         End
         Begin EditLib.fpDoubleSingle ipp_ComVta 
            Height          =   315
            Left            =   2130
            TabIndex        =   15
            Top             =   60
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDoubleSingle ipp_MtoPre 
            Height          =   315
            Left            =   2130
            TabIndex        =   16
            Top             =   420
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpLongInteger ipp_NumCuo 
            Height          =   315
            Left            =   2130
            TabIndex        =   17
            Top             =   1080
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
            MaxValue        =   "0"
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
            Left            =   2130
            TabIndex        =   18
            Top             =   1410
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
            MaxValue        =   "12"
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
         Begin Threed.SSPanel pnl_CuoPBP 
            Height          =   315
            Left            =   10020
            TabIndex        =   19
            Top             =   3120
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CuoIni 
            Height          =   315
            Left            =   2130
            TabIndex        =   20
            Top             =   750
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CosEfe 
            Height          =   315
            Left            =   10020
            TabIndex        =   21
            Top             =   2460
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.0000 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CuoSBP 
            Height          =   315
            Left            =   10020
            TabIndex        =   22
            Top             =   2790
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_PorIni 
            Height          =   315
            Left            =   3600
            TabIndex        =   23
            Top             =   750
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin EditLib.fpDoubleSingle ipp_TasInt 
            Height          =   315
            Left            =   10020
            TabIndex        =   38
            Top             =   60
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDoubleSingle ipp_SegDes 
            Height          =   315
            Left            =   10020
            TabIndex        =   39
            Top             =   390
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            Text            =   "0.000000"
            DecimalPlaces   =   6
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
         Begin EditLib.fpDoubleSingle ipp_SegInm 
            Height          =   315
            Left            =   10020
            TabIndex        =   40
            Top             =   720
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            Text            =   "0.000000"
            DecimalPlaces   =   6
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
         Begin EditLib.fpDoubleSingle ipp_Portes 
            Height          =   315
            Left            =   10020
            TabIndex        =   45
            Top             =   1050
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            Text            =   "0.000000"
            DecimalPlaces   =   6
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
         Begin EditLib.fpDoubleSingle ipp_MtoPBP 
            Height          =   315
            Left            =   2130
            TabIndex        =   46
            Top             =   2400
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpLongInteger ipp_DiaPag 
            Height          =   315
            Left            =   2130
            TabIndex        =   47
            Top             =   2070
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
            MaxValue        =   "28"
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
         Begin EditLib.fpDateTime ipp_DesCli 
            Height          =   315
            Left            =   2130
            TabIndex        =   48
            Top             =   2730
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            ButtonStyle     =   3
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
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime ipp_DesCof 
            Height          =   315
            Left            =   2130
            TabIndex        =   50
            Top             =   3060
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            ButtonStyle     =   3
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
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle ipp_ValAse 
            Height          =   315
            Left            =   2130
            TabIndex        =   52
            Top             =   3390
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDoubleSingle ipp_TasCof 
            Height          =   315
            Left            =   10020
            TabIndex        =   124
            Top             =   1380
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDoubleSingle ipp_ComCof 
            Height          =   315
            Left            =   10020
            TabIndex        =   126
            Top             =   1710
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDoubleSingle ipp_TasMvi 
            Height          =   315
            Left            =   10020
            TabIndex        =   128
            Top             =   2040
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin VB.Label Label15 
            Caption         =   "Tasa Interés Mivivienda:"
            Height          =   315
            Left            =   8010
            TabIndex        =   129
            Top             =   2040
            Width           =   1875
         End
         Begin VB.Label Label14 
            Caption         =   "Tasa Comisión Cofide:"
            Height          =   315
            Left            =   8010
            TabIndex        =   127
            Top             =   1710
            Width           =   1575
         End
         Begin VB.Label Label13 
            Caption         =   "Tasa Interés Cofide:"
            Height          =   315
            Left            =   8010
            TabIndex        =   125
            Top             =   1380
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "Valor Asegurable Inmueb.:"
            Height          =   285
            Left            =   60
            TabIndex        =   53
            Top             =   3390
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "F. Desembolso Cofide:"
            Height          =   315
            Left            =   60
            TabIndex        =   51
            Top             =   3060
            Width           =   1875
         End
         Begin VB.Label Label5 
            Caption         =   "F. Desembolso Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   49
            Top             =   2730
            Width           =   1905
         End
         Begin VB.Label Label3 
            Caption         =   "Portes:"
            Height          =   315
            Left            =   8010
            TabIndex        =   44
            Top             =   1050
            Width           =   1755
         End
         Begin VB.Label lbl_MtoPBP 
            Caption         =   "Monto PBP:"
            Height          =   315
            Left            =   60
            TabIndex        =   37
            Top             =   2400
            Width           =   945
         End
         Begin VB.Label Label27 
            Caption         =   "Monto Solicitado:"
            Height          =   285
            Left            =   60
            TabIndex        =   36
            Top             =   420
            Width           =   1365
         End
         Begin VB.Label Label35 
            Caption         =   "Valor Inmueble:"
            Height          =   285
            Left            =   60
            TabIndex        =   35
            Top             =   60
            Width           =   1575
         End
         Begin VB.Label Label29 
            Caption         =   "Plazo:"
            Height          =   285
            Left            =   60
            TabIndex        =   34
            Top             =   1080
            Width           =   1665
         End
         Begin VB.Label Label25 
            Caption         =   "Período de Gracia:"
            Height          =   285
            Left            =   90
            TabIndex        =   33
            Top             =   1410
            Width           =   1395
         End
         Begin VB.Label Label18 
            Caption         =   "Día de Pago:"
            Height          =   315
            Left            =   60
            TabIndex        =   32
            Top             =   2070
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Seguro Desgrav.:"
            Height          =   315
            Left            =   60
            TabIndex        =   31
            Top             =   1740
            Width           =   1725
         End
         Begin VB.Label lbl_CuoPBP 
            Caption         =   "Cuota con PBP:"
            Height          =   315
            Left            =   8010
            TabIndex        =   30
            Top             =   3120
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "Inicial:"
            Height          =   285
            Left            =   60
            TabIndex        =   29
            Top             =   750
            Width           =   1365
         End
         Begin VB.Label Label26 
            Caption         =   "Tasa Interés:"
            Height          =   315
            Left            =   8010
            TabIndex        =   28
            Top             =   60
            Width           =   1185
         End
         Begin VB.Label Label8 
            Caption         =   "Costo Efectivo:"
            Height          =   315
            Left            =   8010
            TabIndex        =   27
            Top             =   2460
            Width           =   1185
         End
         Begin VB.Label Label10 
            Caption         =   "Seguro Desgravamen:"
            Height          =   315
            Left            =   8010
            TabIndex        =   26
            Top             =   390
            Width           =   1755
         End
         Begin VB.Label Label12 
            Caption         =   "Seguro Inmueble:"
            Height          =   315
            Left            =   8010
            TabIndex        =   25
            Top             =   720
            Width           =   1755
         End
         Begin VB.Label lbl_CuoSBP 
            Caption         =   "Cuota sin PBP:"
            Height          =   315
            Left            =   8010
            TabIndex        =   24
            Top             =   2790
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_Sim_OpeHip_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(cmb_TipMon)
   
   If cmb_Produc.ListIndex > -1 Then
      ipp_SegInm.Value = 0.023
   
      Select Case cmb_Produc.ItemData(cmb_Produc.ListIndex)
         Case 1
            ipp_NumCuo.MinValue = 120
            ipp_NumCuo.MaxValue = 240
            
            ipp_PerGra.MinValue = 0
            ipp_PerGra.MaxValue = 6
            
            ipp_Portes.Value = 2
            
            ipp_MtoPBP.Enabled = True
            
            tab_Cronog.TabVisible(1) = True
            tab_Cronog.TabVisible(2) = True
            tab_Cronog.TabVisible(3) = True
            tab_Cronog.TabVisible(4) = False
            
            tab_Cronog.TabCaption(0) = "Cliente (TNC)"
            tab_Cronog.TabCaption(1) = "Cliente (TC)"
            tab_Cronog.TabCaption(2) = "Mivivienda (TNC)"
            tab_Cronog.TabCaption(3) = "Mivivienda (TC)"
         
         Case 2
            ipp_NumCuo.MinValue = 60
            ipp_NumCuo.MaxValue = 240
            
            ipp_PerGra.MinValue = 0
            ipp_PerGra.MaxValue = 6
            
            ipp_Portes.Value = 2
            
            ipp_MtoPBP.Value = 0
            ipp_MtoPBP.Enabled = False
            
            tab_Cronog.TabVisible(1) = False
            tab_Cronog.TabVisible(2) = False
            tab_Cronog.TabVisible(3) = False
            tab_Cronog.TabVisible(4) = False
            
            tab_Cronog.TabCaption(0) = "Cliente"
            
         Case 3
            ipp_NumCuo.MinValue = 60
            ipp_NumCuo.MaxValue = 240
            
            ipp_PerGra.MinValue = 0
            ipp_PerGra.MaxValue = 6
            
            ipp_Portes.Value = 9
            
            ipp_MtoPBP.Enabled = True
            
            tab_Cronog.TabVisible(1) = True
            tab_Cronog.TabVisible(2) = True
            tab_Cronog.TabVisible(3) = True
            tab_Cronog.TabVisible(4) = True
            
            tab_Cronog.TabCaption(0) = "Cliente (TNC)"
            tab_Cronog.TabCaption(1) = "Cliente (TC)"
            tab_Cronog.TabCaption(2) = "Mivivienda (TNC)"
            tab_Cronog.TabCaption(3) = "Mivivienda (TC)"
            tab_Cronog.TabCaption(4) = "Cofide"
            
         Case 4, 7
            ipp_NumCuo.MinValue = 60
            ipp_NumCuo.MaxValue = 240
            
            ipp_PerGra.MinValue = 0
            ipp_PerGra.MaxValue = 0
            
            ipp_Portes.Value = 9
            
            ipp_MtoPBP.Enabled = True
            
            tab_Cronog.TabVisible(1) = True
            tab_Cronog.TabVisible(2) = True
            tab_Cronog.TabVisible(3) = True
            tab_Cronog.TabVisible(4) = False
            
            tab_Cronog.TabCaption(0) = "Cliente (TNC)"
            tab_Cronog.TabCaption(1) = "Cliente (TC)"
            tab_Cronog.TabCaption(2) = "Cofide (TNC)"
            tab_Cronog.TabCaption(3) = "Cofide (TC)"
            
         Case 9
            ipp_NumCuo.MinValue = 120
            ipp_NumCuo.MaxValue = 180
            
            ipp_PerGra.MinValue = 0
            ipp_PerGra.MaxValue = 0
            
            ipp_Portes.Value = 9
         
            ipp_MtoPBP.Enabled = True
            
            tab_Cronog.TabVisible(1) = True
            tab_Cronog.TabVisible(2) = True
            tab_Cronog.TabVisible(3) = True
            tab_Cronog.TabVisible(4) = False
            
            tab_Cronog.TabCaption(0) = "Cliente (TNC)"
            tab_Cronog.TabCaption(1) = "Cliente (TC)"
            tab_Cronog.TabCaption(2) = "Cofide (TNC)"
            tab_Cronog.TabCaption(3) = "Cofide (TC)"
         
         Case 10
            ipp_NumCuo.MinValue = 120
            ipp_NumCuo.MaxValue = 180
            
            ipp_PerGra.MinValue = 0
            ipp_PerGra.MaxValue = 0
            
            ipp_Portes.Value = 9
         
            ipp_MtoPBP.Enabled = True
            
            tab_Cronog.TabVisible(1) = True
            tab_Cronog.TabVisible(2) = True
            tab_Cronog.TabVisible(3) = True
            tab_Cronog.TabVisible(4) = False
            
            tab_Cronog.TabCaption(0) = "Cliente (TNC)"
            tab_Cronog.TabCaption(1) = "Cliente (TC)"
            tab_Cronog.TabCaption(2) = "Cofide (TNC)"
            tab_Cronog.TabCaption(3) = "Cofide (TC)"
         
         Case 11
            ipp_NumCuo.MinValue = 120
            ipp_NumCuo.MaxValue = 180
            
            ipp_PerGra.MinValue = 0
            ipp_PerGra.MaxValue = 6
            
            ipp_Portes.Value = 9
      
            ipp_MtoPBP.Value = 0
            ipp_MtoPBP.Enabled = False
      
            tab_Cronog.TabVisible(1) = False
            tab_Cronog.TabVisible(2) = False
            tab_Cronog.TabVisible(3) = False
            tab_Cronog.TabVisible(4) = True
            
            tab_Cronog.TabCaption(0) = "Cliente"
            tab_Cronog.TabCaption(4) = "Cofide"
      End Select
   End If
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(ipp_ComVta)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmb_TipSeg_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub cmb_TipSeg_Click()
   Call gs_SetFocus(ipp_DiaPag)
   
   If cmb_TipSeg.ListIndex > -1 Then
      Select Case cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)
         Case 1
            ipp_SegDes.Value = 0.04
         Case 2
            ipp_SegDes.Value = 0.07
      End Select
   End If
End Sub

Private Sub cmb_TipSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipSeg_Click
   End If
End Sub

Private Sub cmd_CalCuo_Click()
   If cmb_Produc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Produc)
      Exit Sub
   End If
   
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If

   If ipp_ComVta.Value = 0 Then
      MsgBox "Debe ingresar Valor de Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta)
      Exit Sub
   End If
   
   If ipp_MtoPre.Value = 0 Then
      MsgBox "Debe ingresar el Monto del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoPre)
      Exit Sub
   End If
   
   If ipp_NumCuo.Value = 0 Then
      MsgBox "Debe ingresar el Número de Cuotas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_NumCuo)
      Exit Sub
   End If
   
   If ipp_DiaPag.Value = 0 Then
      Select Case cmb_Produc.ItemData(cmb_Produc.ListIndex)
         Case 1, 2
            If Not (ipp_DiaPag.Value = 5 Or ipp_DiaPag.Value = 20) Then
               MsgBox "Día de Pago sólo puede ser día 5 o día 20.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_DiaPag)
               Exit Sub
            End If
      
         Case 3, 11
            If Not (ipp_DiaPag.Value = 2 Or ipp_DiaPag.Value = 17) Then
               MsgBox "Día de Pago sólo puede ser día 2 o día 17.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_DiaPag)
               Exit Sub
            End If
      
         Case 4, 7, 9, 10
            If Not (ipp_DiaPag.Value >= 1 And ipp_DiaPag.Value <= 28) Then
               MsgBox "Día de Pago sólo puede estar entre el día 1 y el día 28.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_DiaPag)
               Exit Sub
            End If
      End Select
   End If

   pnl_CuoIni.Caption = Format(CDbl(ipp_ComVta.Text) - CDbl(ipp_MtoPre.Text), "###,##0.00") & " "
   pnl_PorIni.Caption = "(" & Format(CDbl(pnl_CuoIni.Caption) / CDbl(ipp_ComVta.Text) * 100, "##0.00") & "%) "

   Select Case cmb_Produc.ItemData(cmb_Produc.ListIndex)
      Case 2
         Call fs_Cal_miCasita
         
      Case 11
         Call fs_Cal_miCasita_Cofide
   End Select

End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call gs_CentraForm(Me)
   
   Call fs_Inicia
   Call fs_Limpia
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Producto
   cmb_Produc.Clear
   
   cmb_Produc.AddItem "CRC-PBP"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 1
   
   cmb_Produc.AddItem "MICASITA"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 2
   
   cmb_Produc.AddItem "CME"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 3
   
   cmb_Produc.AddItem "MIHOGAR"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 4
   
   cmb_Produc.AddItem "MIVIVIENDA"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 7
   
   cmb_Produc.AddItem "MIVIVIENDA PERUANOS EN EL EXTRANJERO (UNION ANDINA)"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 9
   
   cmb_Produc.AddItem "MIVIVIENDA PERUANOS EN EL EXTRANJERO"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 10
   
   cmb_Produc.AddItem "MICASITA-COFIDE"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 11
   
   'Moneda
   cmb_TipMon.Clear
   
   cmb_TipMon.AddItem "SOLES"
   cmb_TipMon.ItemData(cmb_TipMon.NewIndex) = 1

   cmb_TipMon.AddItem "DOLARES AMERICANOS"
   cmb_TipMon.ItemData(cmb_TipMon.NewIndex) = 2
   
   'Tipo de Seguro de Desgravamen
   cmb_TipSeg.Clear
   
   cmb_TipSeg.AddItem "INDIVIDUAL"
   cmb_TipSeg.ItemData(cmb_TipSeg.NewIndex) = 1

   cmb_TipSeg.AddItem "MANCOMUNADO"
   cmb_TipSeg.ItemData(cmb_TipSeg.NewIndex) = 2

   'Cliente No Concesional
   grd_CliNCo_Listad.ColWidth(0) = 795
   grd_CliNCo_Listad.ColWidth(1) = 1425
   grd_CliNCo_Listad.ColWidth(2) = 1180
   grd_CliNCo_Listad.ColWidth(3) = 1170
   grd_CliNCo_Listad.ColWidth(4) = 1160
   grd_CliNCo_Listad.ColWidth(5) = 1160
   grd_CliNCo_Listad.ColWidth(6) = 1160
   grd_CliNCo_Listad.ColWidth(7) = 1320
   grd_CliNCo_Listad.ColWidth(8) = 1560
   
   grd_CliNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(8) = flexAlignRightCenter

   'Mivivienda No Concesional
   grd_MViNCo_Listad.ColWidth(0) = 695
   grd_MViNCo_Listad.ColWidth(1) = 1415
   grd_MViNCo_Listad.ColWidth(2) = 1070
   grd_MViNCo_Listad.ColWidth(3) = 1070
   grd_MViNCo_Listad.ColWidth(4) = 1080
   grd_MViNCo_Listad.ColWidth(5) = 1080
   grd_MViNCo_Listad.ColWidth(6) = 1080
   grd_MViNCo_Listad.ColWidth(7) = 1080
   grd_MViNCo_Listad.ColWidth(8) = 1080
   grd_MViNCo_Listad.ColWidth(9) = 1290
   
   grd_MViNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(9) = flexAlignRightCenter

   'Mivivienda Concesional
   grd_MViCon_Listad.ColWidth(0) = 770
   grd_MViCon_Listad.ColWidth(1) = 1485
   grd_MViCon_Listad.ColWidth(2) = 1730
   grd_MViCon_Listad.ColWidth(3) = 1740
   grd_MViCon_Listad.ColWidth(4) = 1740
   grd_MViCon_Listad.ColWidth(5) = 1740
   grd_MViCon_Listad.ColWidth(6) = 1740
   
   grd_MViCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MViCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MViCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MViCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MViCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MViCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MViCon_Listad.ColAlignment(6) = flexAlignRightCenter

   'Cliente Concesional
   grd_CliCon_Listad.ColWidth(0) = 770
   grd_CliCon_Listad.ColWidth(1) = 1485
   grd_CliCon_Listad.ColWidth(2) = 2170
   grd_CliCon_Listad.ColWidth(3) = 2160
   grd_CliCon_Listad.ColWidth(4) = 2170
   grd_CliCon_Listad.ColWidth(5) = 2170
   
   grd_CliCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(5) = flexAlignRightCenter

   'Cofide No Concesional
   grd_CofNCo_Listad.ColWidth(0) = 770
   grd_CofNCo_Listad.ColWidth(1) = 1485
   grd_CofNCo_Listad.ColWidth(2) = 1730
   grd_CofNCo_Listad.ColWidth(3) = 1740
   grd_CofNCo_Listad.ColWidth(4) = 1740
   grd_CofNCo_Listad.ColWidth(5) = 1740
   grd_CofNCo_Listad.ColWidth(6) = 1740
   
   grd_CofNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(6) = flexAlignRightCenter
End Sub

Private Sub fs_Limpia()
   cmb_Produc.ListIndex = -1
   cmb_TipMon.ListIndex = -1
   
   ipp_TipCam.Value = 0
   
   ipp_ComVta.Value = 0
   ipp_MtoPre.Value = 0
   
   pnl_CuoIni.Caption = "0.00 "
   pnl_PorIni.Caption = "(0.00%) "
   
   ipp_NumCuo.Value = 0
   ipp_PerGra.Value = 0
   
   cmb_TipSeg.ListIndex = -1
   
   ipp_DiaPag.Value = 0
   ipp_MtoPBP.Value = 0
   
   ipp_DesCli.Text = Format(date, "dd/mm/yyyy")
   ipp_DesCof.Text = Format(date, "dd/mm/yyyy")
   
   ipp_ValAse.Value = 0
   
   ipp_TasInt.Value = 0
   ipp_SegDes.Value = 0
   ipp_SegInm.Value = 0
   ipp_Portes.Value = 0
   
   pnl_CosEfe.Caption = "0.0000 "
   pnl_CuoSBP.Caption = "0.00 "
   pnl_CuoPBP.Caption = "0.00 "
   
   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   Call gs_LimpiaGrid(grd_MViCon_Listad)
   Call gs_LimpiaGrid(grd_CofNCo_Listad)
   
   pnl_CliNCo_Capita.Caption = "0.00 "
   pnl_CliNCo_Intere.Caption = "0.00 "
   pnl_CliNCo_SegPre.Caption = "0.00 "
   pnl_CliNCo_SegViv.Caption = "0.00 "
   pnl_CliNCo_OtrCar.Caption = "0.00 "
   pnl_CliNCo_TotCuo.Caption = "0.00 "
   
   pnl_CliCon_Capita.Caption = "0.00 "
   pnl_CliCon_Intere.Caption = "0.00 "
   pnl_CliCon_TotCuo.Caption = "0.00 "
   
   pnl_MViNCo_Capita.Caption = "0.00 "
   pnl_MViNCo_Intere.Caption = "0.00 "
   pnl_MViNCo_SegPre.Caption = "0.00 "
   pnl_MViNCo_SegViv.Caption = "0.00 "
   pnl_MViNCo_OtrCar.Caption = "0.00 "
   pnl_MViNCo_TotCuo.Caption = "0.00 "
   
   pnl_MViCon_Capita.Caption = "0.00 "
   pnl_MViCon_Intere.Caption = "0.00 "
   pnl_MViCon_Comisi.Caption = "0.00 "
   pnl_MViCon_TotCuo.Caption = "0.00 "
   
   pnl_CofNCo_Capita.Caption = "0.00 "
   pnl_CofNCo_Intere.Caption = "0.00 "
   pnl_CofNCo_Comisi.Caption = "0.00 "
   pnl_CofNCo_TotCuo.Caption = "0.00 "
End Sub

Private Sub ipp_ComCof_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_ComCof_KeyPress(KeyAscii As Integer)
      If ipp_TasMvi.Enabled Then
         Call gs_SetFocus(ipp_TasMvi)
      Else
         Call gs_SetFocus(cmd_CalCuo)
      End If
End Sub

Private Sub ipp_ComVta_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_ComVta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoPre)
   End If
End Sub

Private Sub ipp_ComVta_LostFocus()
   ipp_ValAse.Value = ipp_ComVta.Value * 0.72
End Sub

Private Sub fs_Limpia_Calcul()
   pnl_CuoIni.Caption = "0.00 "
   pnl_PorIni.Caption = "0.00% "
   
   pnl_CosEfe.Caption = "0.00 "
   
   pnl_CuoSBP.Caption = "0.00 "
   pnl_CuoPBP.Caption = "0.00 "

   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   Call gs_LimpiaGrid(grd_MViCon_Listad)
   Call gs_LimpiaGrid(grd_CofNCo_Listad)
   
   pnl_CliNCo_Capita.Caption = "0.00 "
   pnl_CliNCo_Intere.Caption = "0.00 "
   pnl_CliNCo_SegPre.Caption = "0.00 "
   pnl_CliNCo_SegViv.Caption = "0.00 "
   pnl_CliNCo_OtrCar.Caption = "0.00 "
   pnl_CliNCo_TotCuo.Caption = "0.00 "
   
   pnl_CliCon_Capita.Caption = "0.00 "
   pnl_CliCon_Intere.Caption = "0.00 "
   pnl_CliCon_TotCuo.Caption = "0.00 "
   
   pnl_MViNCo_Capita.Caption = "0.00 "
   pnl_MViNCo_Intere.Caption = "0.00 "
   pnl_MViNCo_SegPre.Caption = "0.00 "
   pnl_MViNCo_SegViv.Caption = "0.00 "
   pnl_MViNCo_OtrCar.Caption = "0.00 "
   pnl_MViNCo_TotCuo.Caption = "0.00 "
   
   pnl_MViCon_Capita.Caption = "0.00 "
   pnl_MViCon_Intere.Caption = "0.00 "
   pnl_MViCon_Comisi.Caption = "0.00 "
   pnl_MViCon_TotCuo.Caption = "0.00 "
   
   pnl_CofNCo_Capita.Caption = "0.00 "
   pnl_CofNCo_Intere.Caption = "0.00 "
   pnl_CofNCo_Comisi.Caption = "0.00 "
   pnl_CofNCo_TotCuo.Caption = "0.00 "
End Sub

Private Sub ipp_DesCli_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_DesCli_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_DesCof.Enabled Then
         Call gs_SetFocus(ipp_DesCof)
      Else
         Call gs_SetFocus(ipp_ValAse)
      End If
   End If
End Sub

Private Sub ipp_DesCof_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_DiaPag_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_DiaPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_MtoPBP.Enabled Then
         Call gs_SetFocus(ipp_MtoPBP)
      Else
         Call gs_SetFocus(ipp_DesCli)
      End If
   End If
End Sub

Private Sub ipp_MtoPBP_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_MtoPre_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_MtoPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumCuo)
   End If
End Sub

Private Sub ipp_NumCuo_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_NumCuo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerGra)
   End If
End Sub

Private Sub ipp_PerGra_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_PerGra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipSeg)
   End If
End Sub

Private Sub ipp_Portes_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_Portes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_TasCof.Enabled Then
         Call gs_SetFocus(ipp_TasCof)
      ElseIf ipp_ComCof.Enabled Then
         Call gs_SetFocus(ipp_ComCof)
      ElseIf ipp_TasMvi.Enabled Then
         Call gs_SetFocus(ipp_TasMvi)
      Else
         Call gs_SetFocus(cmd_CalCuo)
      End If
   End If
End Sub

Private Sub ipp_SegDes_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_SegDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SegInm)
   End If
End Sub

Private Sub ipp_SegInm_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_SegInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Portes)
   End If
End Sub

Private Sub ipp_TasCof_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_TasCof_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_ComCof.Enabled Then
         Call gs_SetFocus(ipp_ComCof)
      ElseIf ipp_TasMvi.Enabled Then
         Call gs_SetFocus(ipp_TasMvi)
      Else
         Call gs_SetFocus(cmd_CalCuo)
      End If
   End If
End Sub

Private Sub ipp_TasInt_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_TasInt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SegDes)
   End If
End Sub

Private Sub ipp_TasMvi_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_TasMvi_KeyPress(KeyAscii As Integer)
   Call gs_SetFocus(cmd_CalCuo)
End Sub

Private Sub ipp_ValAse_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_ValAse_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TasInt)
   End If
End Sub

Private Sub fs_Cal_miCasita()
   Dim r_arr_CliNco()      As modcal_g_est_CuoCli

   Call gs_Cronog_MiCasita(r_arr_CliNco(), ipp_ValAse.Value, ipp_MtoPre.Value, ipp_NumCuo, 2, ipp_TasInt.Value, ipp_SegDes.Value, 1, ipp_SegInm.Value, ipp_Portes.Value, ipp_DesCli.Text, ipp_DiaPag.Value, ipp_PerGra.Value, , 1)
   Call fs_Cronog_CliNco(r_arr_CliNco())
   
   'Obteniendo Cuota Mensual
   pnl_CuoSBP.Caption = Format(r_arr_CliNco(2).CuoCli_ValCuo, "###,###,##0.00") & " "
   pnl_CuoPBP.Caption = Format(r_arr_CliNco(2).CuoCli_ValCuo, "###,###,##0.00") & " "
      
   pnl_CosEfe.Caption = Format(gf_Cronog_CosEfe(r_arr_CliNco(), ipp_TasInt.Value, ipp_MtoPre.Value), "###,##0.00") & " "
   
End Sub

Private Sub fs_Cal_miCasita_Cofide()
   Dim r_arr_CliNco()      As modcal_g_est_CuoCli

   Call gs_Cronog_MiCasita(r_arr_CliNco(), ipp_ValAse.Value, ipp_MtoPre.Value, ipp_NumCuo, 2, ipp_TasInt.Value, ipp_SegDes.Value, 1, ipp_SegInm.Value, ipp_Portes.Value, ipp_DesCli.Text, ipp_DiaPag.Value, ipp_PerGra.Value, , 1)
   Call fs_Cronog_CliNco(r_arr_CliNco())
   
   'Obteniendo Cuota Mensual
   pnl_CuoSBP.Caption = Format(r_arr_CliNco(2).CuoCli_ValCuo, "###,###,##0.00") & " "
   pnl_CuoPBP.Caption = Format(r_arr_CliNco(2).CuoCli_ValCuo, "###,###,##0.00") & " "
      
   pnl_CosEfe.Caption = Format(gf_Cronog_CosEfe(r_arr_CliNco(), ipp_TasInt.Value, ipp_MtoPre.Value), "###,##0.00") & " "
   
   Call gs_Cronog_miCasitaCOFIDE_Cof(r_arr_CliNco(), ipp_MtoPre.Value, 0, ipp_NumCuo.Value, ipp_PerGra.Value, ipp_TasCof, ipp_ComCof, ipp_DesCof.Text, 4)
   Call fs_Cronog_MviCofNco(r_arr_CliNco())
End Sub

Private Sub fs_Cronog_CliNco(p_Arregl() As modcal_g_est_CuoCli)
   Dim r_dbl_Cuo_Capita    As Double
   Dim r_dbl_Cuo_Intere    As Double
   Dim r_dbl_Cuo_SegPre    As Double
   Dim r_dbl_Cuo_SegViv    As Double
   Dim r_dbl_Cuo_Portes    As Double
   Dim r_dbl_Cuo_TotCuo    As Double
   
   Dim r_dbl_Tot_Capita    As Double
   Dim r_dbl_Tot_Intere    As Double
   Dim r_dbl_Tot_SegPre    As Double
   Dim r_dbl_Tot_SegViv    As Double
   Dim r_dbl_Tot_Portes    As Double
   Dim r_dbl_Tot_TotCuo    As Double
   
   Dim r_int_Contad        As Integer
   
   'Mostrando Cronograma
   grd_CliNCo_Listad.Redraw = False
   
   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   
   r_dbl_Tot_Capita = 0
   r_dbl_Tot_Intere = 0
   r_dbl_Tot_SegPre = 0
   r_dbl_Tot_SegViv = 0
   r_dbl_Tot_Portes = 0
   r_dbl_Tot_TotCuo = 0
   
   For r_int_Contad = 1 To UBound(p_Arregl)
      grd_CliNCo_Listad.Rows = grd_CliNCo_Listad.Rows + 1
      grd_CliNCo_Listad.Row = grd_CliNCo_Listad.Rows - 1
      
      
      r_dbl_Cuo_Capita = CDbl(Format(p_Arregl(r_int_Contad).CuoCli_Capita, "###,##0.00"))
      r_dbl_Cuo_Intere = CDbl(Format(p_Arregl(r_int_Contad).CuoCli_Intere, "###,##0.00"))
      r_dbl_Cuo_SegPre = CDbl(Format(p_Arregl(r_int_Contad).CuoCli_SegPre, "###,##0.00"))
      r_dbl_Cuo_SegViv = CDbl(Format(p_Arregl(r_int_Contad).CuoCli_SegViv, "###,##0.00"))
      r_dbl_Cuo_Portes = CDbl(Format(p_Arregl(r_int_Contad).CuoCli_Portes, "###,##0.00"))
      r_dbl_Cuo_TotCuo = CDbl(Format(r_dbl_Cuo_Capita + r_dbl_Cuo_Intere + r_dbl_Cuo_SegPre + r_dbl_Cuo_SegViv + r_dbl_Cuo_Portes, "###,##0.00"))
      
      r_dbl_Tot_Capita = r_dbl_Tot_Capita + r_dbl_Cuo_Capita
      r_dbl_Tot_Intere = r_dbl_Tot_Intere + r_dbl_Cuo_Intere
      r_dbl_Tot_SegPre = r_dbl_Tot_SegPre + r_dbl_Cuo_SegPre
      r_dbl_Tot_SegViv = r_dbl_Tot_SegViv + r_dbl_Cuo_SegViv
      r_dbl_Tot_Portes = r_dbl_Tot_Portes + r_dbl_Cuo_Portes
      r_dbl_Tot_TotCuo = r_dbl_Tot_TotCuo + r_dbl_Cuo_TotCuo
      
      grd_CliNCo_Listad.Col = 0
      grd_CliNCo_Listad.Text = Format(r_int_Contad, "000")
      
      grd_CliNCo_Listad.Col = 1
      grd_CliNCo_Listad.Text = p_Arregl(r_int_Contad).CuoCli_FecVct
   
      grd_CliNCo_Listad.Col = 2
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_Capita, "###,##0.00")
   
      grd_CliNCo_Listad.Col = 3
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_Intere, "###,##0.00")
   
      grd_CliNCo_Listad.Col = 4
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_SegPre, "###,##0.00")
   
      grd_CliNCo_Listad.Col = 5
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_SegViv, "###,##0.00")
   
      grd_CliNCo_Listad.Col = 6
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_Portes, "###,##0.00")
   
      grd_CliNCo_Listad.Col = 7
      grd_CliNCo_Listad.Text = Format(r_dbl_Cuo_TotCuo, "###,##0.00")
      
      grd_CliNCo_Listad.Col = 8
      grd_CliNCo_Listad.Text = Format(p_Arregl(r_int_Contad).CuoCli_SalCap, "###,##0.00")
   Next r_int_Contad
   
   grd_CliNCo_Listad.Redraw = True
   
   Call gs_UbiIniGrid(grd_CliNCo_Listad)
            
   pnl_CliNCo_Capita.Caption = Format(r_dbl_Tot_Capita, "###,##0.00") & " "
   pnl_CliNCo_Intere.Caption = Format(r_dbl_Tot_Intere, "###,##0.00") & " "
   pnl_CliNCo_SegPre.Caption = Format(r_dbl_Tot_SegPre, "###,##0.00") & " "
   pnl_CliNCo_SegViv.Caption = Format(r_dbl_Tot_SegViv, "###,##0.00") & " "
   pnl_CliNCo_OtrCar.Caption = Format(r_dbl_Tot_Portes, "###,##0.00") & " "
   pnl_CliNCo_TotCuo.Caption = Format(r_dbl_Tot_TotCuo, "###,##0.00") & " "
End Sub

Private Sub fs_Cronog_MviCofNco(p_Arregl() As modcal_g_est_CuoCli)
   Dim r_dbl_Capita        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
   Dim r_dbl_TotCuo        As Double
   
   Dim r_int_Contad        As Integer
   
   'Cofide No Concesional
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_Comisi = 0
   r_dbl_TotCuo = 0
   
   grd_CofNCo_Listad.Redraw = False
   For r_int_Contad = 1 To UBound(p_Arregl)
      grd_CofNCo_Listad.Rows = grd_CofNCo_Listad.Rows + 1
      grd_CofNCo_Listad.Row = grd_CofNCo_Listad.Rows - 1
      
      'Número de Cuota
      grd_CofNCo_Listad.Col = 0
      grd_CofNCo_Listad.Text = Format(r_int_Contad, "000")
   
      'Fecha de Vencimiento
      grd_CofNCo_Listad.Col = 1
      grd_CofNCo_Listad.Text = p_Arregl(r_int_Contad).CuoCli_FecVct
   
      'Capital
      grd_CofNCo_Listad.Col = 2
      grd_CofNCo_Listad.Text = Format(p_Arregl(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
      r_dbl_Capita = r_dbl_Capita + CDbl(grd_CofNCo_Listad)
      
      'Interes
      grd_CofNCo_Listad.Col = 3
      grd_CofNCo_Listad.Text = Format(p_Arregl(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
      r_dbl_Intere = r_dbl_Intere + CDbl(grd_CofNCo_Listad)
   
      'Comisión
      grd_CofNCo_Listad.Col = 4
      grd_CofNCo_Listad.Text = Format(p_Arregl(r_int_Contad).CuoCli_Comisi, "###,###,##0.00")
      r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_CofNCo_Listad)
   
      'Valor Cuota
      grd_CofNCo_Listad.Col = 5
      grd_CofNCo_Listad.Text = Format(p_Arregl(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
      r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CofNCo_Listad)
   
      'Saldo Capital
      grd_CofNCo_Listad.Col = 6
      grd_CofNCo_Listad.Text = Format(p_Arregl(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
   Next r_int_Contad
   
   grd_CofNCo_Listad.Redraw = True
   
   pnl_CofNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CofNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CofNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
   pnl_CofNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub


