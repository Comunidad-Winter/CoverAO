VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "Cswsk32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "CoverAO 1.0"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   315
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H00000000&
      Height          =   2400
      Left            =   9015
      ScaleHeight     =   158.025
      ScaleMode       =   0  'User
      ScaleWidth      =   161
      TabIndex        =   27
      Top             =   2220
      Width           =   2415
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   10185
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   7
      Top             =   7350
      Width           =   1455
      Begin VB.PictureBox Shape2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H001274FF&
         BorderStyle     =   0  'None
         FillColor       =   &H001274FF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H001274FF&
         Height          =   45
         Index           =   5
         Left            =   0
         ScaleHeight     =   45
         ScaleWidth      =   45
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.PictureBox Shape2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H001274FF&
         BorderStyle     =   0  'None
         FillColor       =   &H001274FF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H001274FF&
         Height          =   45
         Index           =   4
         Left            =   0
         ScaleHeight     =   45
         ScaleWidth      =   45
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.PictureBox Shape2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H001274FF&
         BorderStyle     =   0  'None
         FillColor       =   &H001274FF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H001274FF&
         Height          =   45
         Index           =   3
         Left            =   0
         ScaleHeight     =   45
         ScaleWidth      =   45
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.PictureBox Shape2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H001274FF&
         BorderStyle     =   0  'None
         FillColor       =   &H001274FF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H001274FF&
         Height          =   45
         Index           =   2
         Left            =   0
         ScaleHeight     =   45
         ScaleWidth      =   45
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.PictureBox Shape2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H001274FF&
         BorderStyle     =   0  'None
         FillColor       =   &H001274FF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   1
         Left            =   720
         ScaleHeight     =   45
         ScaleWidth      =   45
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001274FF&
         Height          =   210
         Index           =   4
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001274FF&
         Height          =   210
         Index           =   3
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001274FF&
         Height          =   210
         Index           =   2
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001274FF&
         Height          =   210
         Index           =   5
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001274FF&
         Height          =   210
         Index           =   1
         Left            =   870
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape UserP 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   45
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   11
      Left            =   6090
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   10
         Left            =   0
         TabIndex        =   39
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   10
      Left            =   5505
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   9
         Left            =   0
         TabIndex        =   38
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   9
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   8
         Left            =   0
         TabIndex        =   37
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   8
      Left            =   4335
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   7
         Left            =   0
         TabIndex        =   36
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   7
      Left            =   3750
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10000"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   6
         Left            =   0
         TabIndex        =   35
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   6
      Left            =   3165
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   5
         Left            =   0
         TabIndex        =   34
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   5
      Left            =   2580
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   4
         Left            =   0
         TabIndex        =   33
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   1410
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   2
         Left            =   0
         TabIndex        =   31
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   1995
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   3
         Left            =   0
         TabIndex        =   32
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   825
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   1
         Left            =   0
         TabIndex        =   30
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Index           =   0
         Left            =   0
         TabIndex        =   29
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   210
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   6
      Top             =   2055
      Width           =   8160
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      IntegralHeight  =   0   'False
      Left            =   8850
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   210
      MaxLength       =   500
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1755
      Visible         =   0   'False
      Width           =   7470
   End
   Begin VB.TextBox SendCMSTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   210
      MaxLength       =   500
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1755
      Visible         =   0   'False
      Width           =   7470
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1455
      Left            =   240
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   180
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image nuevocorreo 
      Height          =   255
      Left            =   9645
      Picture         =   "frmMain.frx":0089
      ToolTipText     =   "Nuevo correo"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   1
      Left            =   8820
      TabIndex        =   28
      Top             =   870
      Width           =   1815
   End
   Begin VB.Image imgHora 
      Height          =   480
      Left            =   6673
      Top             =   8430
      Width           =   1695
   End
   Begin VB.Image cmdMensaje 
      Height          =   255
      Left            =   7800
      Top             =   1725
      Width           =   555
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   2
      Left            =   10740
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   1
      Left            =   9660
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   0
      Left            =   8580
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   5
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4935
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   4
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4350
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   3
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3765
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   2
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3180
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   1
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   2595
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   0
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   2010
      Width           =   1890
   End
   Begin VB.Image cmdHechizos 
      Height          =   420
      Index           =   3
      Left            =   11460
      MousePointer    =   99  'Custom
      Top             =   3405
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdHechizos 
      Height          =   420
      Index           =   2
      Left            =   11475
      MousePointer    =   99  'Custom
      Top             =   2910
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdCerrar 
      Height          =   225
      Left            =   11580
      Top             =   180
      Width           =   255
   End
   Begin VB.Label lblDext 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   25
      Top             =   8550
      Width           =   345
   End
   Begin VB.Label lblStrg 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   24
      Top             =   8340
      Width           =   345
   End
   Begin VB.Image modoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":34C8
      ToolTipText     =   "Seguro"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image nomodoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":3906
      ToolTipText     =   "Seguro"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image cmdHechizos 
      Height          =   390
      Index           =   0
      Left            =   8775
      MouseIcon       =   "frmMain.frx":3D44
      MousePointer    =   99  'Custom
      Top             =   4935
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image cmdHechizos 
      Height          =   390
      Index           =   1
      Left            =   10650
      MouseIcon       =   "frmMain.frx":3E96
      MousePointer    =   99  'Custom
      Top             =   4935
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   12
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10320
      TabIndex        =   11
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10320
      TabIndex        =   10
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   9
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   8
      Top             =   5850
      Width           =   1350
   End
   Begin VB.Image cmdMinimizar 
      Height          =   225
      Left            =   11280
      Top             =   180
      Width           =   225
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   5
      Top             =   5745
      Width           =   1110
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10950
      TabIndex        =   4
      Top             =   870
      Width           =   435
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NickDelPersonaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8610
      TabIndex        =   3
      Top             =   180
      Width           =   2625
   End
   Begin VB.Label lblInvInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   9000
      TabIndex        =   2
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Shape shpvida 
      BackColor       =   &H00000080&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8745
      Top             =   5880
      Width           =   1365
   End
   Begin VB.Shape shpEnergia 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   135
      Left            =   8745
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape shpmana 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8745
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape shpHambre 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   135
      Left            =   10320
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape shpSed 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      Height          =   135
      Left            =   10320
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Image lblDropGold 
      Height          =   300
      Left            =   10260
      Top             =   5670
      Width           =   300
   End
   Begin VB.Image nomodocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":3FE8
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image modocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":4426
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMiniCerra 
      Enabled         =   0   'False
      Height          =   315
      Left            =   11340
      Top             =   150
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa desconocido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   254
      Index           =   0
      Left            =   8640
      TabIndex        =   1
      Top             =   7035
      Width           =   3097
   End
   Begin VB.Image InvEqu 
      Height          =   4275
      Left            =   8580
      Top             =   1230
      Width           =   3240
   End
   Begin VB.Shape ExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8820
      Top             =   900
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
' soporte@coverao.com.ar
'   - Relase Number 1
'*****************************************************************

Option Explicit
 
Public MouseX As Long
Public MouseY As Long


'Barrin
Dim UltimoIndex As Integer

Public UltPos As Integer
Public UltPosInterface As Integer
Public UltPosSolapas As Integer
Public CentroActual As Byte
'

Private m_Jpeg As clsJpeg
Private m_FileName As String

'Configuracion teclas
Public MouseBoton As Long, MouseShift As Long

Private Valor As Byte
Public IsPlaying As Byte
 
Private Const EM_GETLINE = &HC4, EM_LINELENGTH = &HC1




Private Sub cmdMensaje_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdMensaje.Picture = General_Load_Skin_Picture_From_Resource_Ex("modotextodown")
End Sub

Private Sub cmdMensaje_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMensaje.Picture = General_Load_Skin_Picture_From_Resource_Ex("modotextoover")
End Sub

Private Sub cmdMensaje_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call cmdMensaje_MouseMove(Button, Shift, X, Y)
frmMensaje.PopupMenuMensaje
cmdMensaje.Picture = General_Load_Skin_Picture_From_Resource_Ex("modotextoover")
End Sub

Private Sub cmdMinimizar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 imgMiniCerra.Picture = General_Load_Skin_Picture_From_Resource_Ex("minimizardown")
End Sub

Private Sub cmdMinimizar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 imgMiniCerra.Picture = General_Load_Skin_Picture_From_Resource_Ex("minimizarover")
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 imgMiniCerra.Picture = General_Load_Skin_Picture_From_Resource_Ex("cerrarover")
End Sub

Private Sub cmdMinimizar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.WindowState = vbMinimized
imgMiniCerra.Picture = Nothing
End Sub

Private Sub cmdCerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Audio.PlayWave(SND_CLICK)

 imgMiniCerra.Picture = General_Load_Skin_Picture_From_Resource_Ex("cerrardown")
End Sub

Private Sub cmdCerrar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CerrarJuego
End Sub

Private Sub Command1_Click()
Call WriteSOSShowList
     '       frmPregunta.SetAccion 8, "asd"

       '     If frmMain.Visible Then frmPregunta.Show , frmMain
End Sub

Private Sub Form_Activate()
    If SendTxt.Visible Then SendTxt.SetFocus
End Sub



Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim map_x As Byte
Dim map_y As Byte

Call Char_MapPosGet(CurrentUser.UserCharIndex, map_x, map_y)


If UltPos <> Index Then
    
    If UltPos >= 0 Then
        If Index = 1 Then
            Label2(Index).Caption = CurrentUser.UserPercExp & "%"
            
        Else
            If VerLugar = 1 Then
                Label2(Index).Caption = Map_Name_Get
            Else
                Label2(Index).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.UserMap & ", " & map_x & ", " & map_y
            End If
            
        End If
    End If
    
    If Index = 1 Then
        Label2(Index).Caption = CurrentUser.UserExp & "/" & CurrentUser.UserPasarNivel
    Else
        If VerLugar = 1 Then
            Label2(Index).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.UserMap & ", " & map_x & ", " & map_y
        Else
            Label2(Index).Caption = Map_Name_Get
        End If
    End If
    
    If CurrentUser.UserPasarNivel = 0 Then
        Label2(1).Caption = Locale_GUI_Frase(173)
    End If
    
    UltPos = Index
End If

End Sub


Private Sub lblDropGold_Click()

   Inventario.SelectGold

If Not Comerciando Then
    If CurrentUser.UserGLD > 0 Then
        frmCantidad.Show vbModeless, frmMain
    End If
Else
    Call AddtoRichTextBox(Locale_GUI_Frase(236), 255, 0, 32, False, False, False)
End If

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hwnd)
Call Auto_Drag(Me.hwnd)
End Sub
Private Sub Form_Unload(Cancel As Integer)
StopURLDetect
End Sub
Private Sub hlst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Nuevo
If CurrentUser.UsingSkill = magia Then
    Call FormParser.Parse_Form(frmMain)
    CurrentUser.UsingSkill = 0
    Call WriteCastSpell(0)
End If

End Sub

Private Sub hlst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub imgHora_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgHora.ToolTipText = Locale_GUI_Frase(302) & " " & Get_Time_String
End Sub
Private Sub imgHora_MouseUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call AddtoRichTextBox(Locale_GUI_Frase(302) & " " & Get_Time_String, 0, 0, 0, 0, 0, 0, 4)
End Sub
Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub modocombate_Click()
Call WriteCombatModeToggle
IScombate = Not IScombate
Call Mod_General.modocombate
End Sub
Public Sub DibujarSeguro()
modoseguro.Visible = True
nomodoseguro.Visible = False
End Sub

Private Sub nomodocombate_Click()
Call WriteCombatModeToggle
IScombate = Not IScombate
Call Mod_General.modocombate
End Sub
Public Sub DesDibujarSeguro()
modoseguro.Visible = False
nomodoseguro.Visible = True
End Sub

Private Sub picInv_Paint()
    RenderInv = True
End Sub
Private Sub picMacro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
MacroIndex = Index
If FrmBindKey.Visible = True Then Exit Sub
    
    If ClientTCP.DeadCheck Then Exit Sub
    
    If Button = vbKeyRButton Or Button = vbRightButton Then
        MacroIndex = Index
        FrmBindKey.Caption = Locale_GUI_Frase(205) & ": F" & Index
        FrmBindKey.Show vbModeless, frmMain
    Else
       Call UsarMacro(Index)
    End If
End Sub
Private Sub picMacro_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Nuevo
If UltimoIndex <> Index Then
    'If UltimoIndex >= 0 Then DibujarMenuMacros UltimoIndex + 1
    'DibujarMenuMacros Index + 1, 1
    UltimoIndex = Index
End If

End Sub

Private Sub picMacro_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If FrmBindKey.Visible = True Or ClientTCP.DeadCheck Then Exit Sub
End Sub

Private Sub lblMacro_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call picMacro_KeyUp(Index, KeyCode, Shift)
End Sub
Private Sub lblMacro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call picMacro_MouseUp(Index + 1, Button, Shift, X, Y)
End Sub
Private Sub TirarItem()

    If ClientTCP.DeadCheck Or Comerciando Then Exit Sub
     
     If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        frmCantidad.Show vbModeless, frmMain
    End If

End Sub
 
Private Sub AgarrarItem()

    If ClientTCP.DeadCheck Or Comerciando Then Exit Sub
    
    Call WritePickUp
End Sub

 Private Sub Form_Load()
    On Error Resume Next
    Call StartURLDetect(RecTxt.hwnd, Me.hwnd)
    
    Me.Picture = General_Load_Skin_Picture_From_Resource_Ex("todo")
    Me.Caption = Form_Caption
    Call Make_Transparent_Richtext(RecTxt.hwnd)
    Call CambiaCentro(CentroInventario)
   
    UltPos = -1
    UltimoIndex = -1
    UltPosInterface = -1
    UltPosSolapas = -1
    
    Call FormParser.Parse_Form(Me)

    Valor = 0 'Amigos
    'Me.Left = 0
    'Me.Top = 0
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseX = X
    MouseY = Y

    Dim map_x As Byte, map_y As Byte
    
    If UltimoIndex >= 0 Then
        'DibujarMenuMacros UltimoIndex + 1
        UltimoIndex = -1
    End If
    
    If UltPos >= 0 Then
        Call Char_MapPosGet(CurrentUser.UserCharIndex, map_x, map_y)
        
        If UltPos = 1 Then
            Label2(UltPos).Caption = CurrentUser.UserPercExp & "%"
        Else
            If VerLugar = 1 Then
                Label2(UltPos).Caption = Map_Name_Get
            Else
                Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.UserMap & ", " & map_x & ", " & map_y
            End If
        End If
        
        If CurrentUser.UserPasarNivel = 0 Then
            Label2(1).Caption = Locale_GUI_Frase(173)
        End If
        
        UltPos = -1
        
    End If
    

    Call RestaurarCentroActual
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub MostrarCentroInventario()

    InvEqu.Picture = General_Load_Skin_Picture_From_Resource_Ex("centroinventario")
    picInv.Visible = True
    lblInvInfo.Visible = True
    lblInvInfo = vbNullString
    CurrentUser.LastItem = 0
    
End Sub
Private Sub OcultarCentroInventario()
    picInv.Visible = False
    lblInvInfo.Visible = False
    CurrentUser.LastItem = 0
 
End Sub

Private Sub MostrarCentroHechizos()
    InvEqu.Picture = General_Load_Skin_Picture_From_Resource_Ex("centrohechizos")
    cmdHechizos(0).Visible = True
    cmdHechizos(1).Visible = True
    cmdHechizos(2).Visible = True
    cmdHechizos(3).Visible = True
    hlst.Visible = True
End Sub

Private Sub OcultarCentroHechizos()
    hlst.Visible = False
    cmdHechizos(0).Visible = False
    cmdHechizos(1).Visible = False
    cmdHechizos(2).Visible = False
    cmdHechizos(3).Visible = False
End Sub

Private Sub MostrarCentroMenu()
    cmdMenu(0).Visible = True
    cmdMenu(1).Visible = True
    cmdMenu(2).Visible = True
    cmdMenu(3).Visible = True
    cmdMenu(4).Visible = True
    cmdMenu(5).Visible = True
    InvEqu.Picture = General_Load_Skin_Picture_From_Resource_Ex("centromenu")
End Sub

Private Sub OcultarCentroMenu()
    cmdMenu(0).Visible = False
    cmdMenu(1).Visible = False
    cmdMenu(2).Visible = False
    cmdMenu(3).Visible = False
    cmdMenu(4).Visible = False
    cmdMenu(5).Visible = False
End Sub

Public Sub CambiaCentro(NuevoCentro As Byte)

CentroActual = NuevoCentro

If NuevoCentro = CentroMenu Then
    Call MostrarCentroMenu
    Call OcultarCentroHechizos
    Call OcultarCentroInventario
ElseIf NuevoCentro = CentroHechizos Then
    Call MostrarCentroHechizos
    Call OcultarCentroMenu
    Call OcultarCentroInventario
Else
    Call MostrarCentroInventario
    Call OcultarCentroHechizos
    Call OcultarCentroMenu
End If

End Sub

Private Sub picInv_DblClick()

    'If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
     
    Call UsarItem
 

End Sub
 Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub
Private Sub SendTxt_Change()
stxtbuffer = SendTxt.Text
End Sub
Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Or (KeyAscii = 126) Or (KeyAscii = 176) Then _
        KeyAscii = 0
End Sub
Private Sub CompletarEnvioMensajes()

Select Case CurrentUser.SendingType
    Case 1
        SendTxt.Text = vbNullString
    Case 2
        SendTxt.Text = "-"
    Case 3
        SendTxt.Text = ("\" & CurrentUser.sndPrivateTo & " ")
    Case 4
        SendTxt.Text = "/CMSG "
    Case 5
        SendTxt.Text = "/GMSG "
    Case 6
        SendTxt.Text = "/GRMG "
    Case 7
        SendTxt.Text = ";"
    Case 8
        SendTxt.Text = "/FMMG "
End Select

stxtbuffer = SendTxt.Text
SendTxt.SelStart = Len(SendTxt.Text)

End Sub
Private Sub Enviar_SendTxt()
    
    
    On Error GoTo Enviar_SendTxt_Err
    
    Dim str1 As String
    Dim str2 As String
    
    If Len(stxtbuffer) > 255 Then stxtbuffer = mid$(stxtbuffer, 1, 255)
        
    If Len(Trim(stxtbuffer)) <= 80 Then
    
        Select Case Left$(stxtbuffer, 1)
        
        Case "/" 'Send text
            Call ClientTCP.ParseUserCommand(stxtbuffer)
        
        Case "-" ''Shout
            If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> vbNullString Then
            
                Dim MsjGritar As String
                MsjGritar = Trim(mid(stxtbuffer, InStr(stxtbuffer, "-") + 1))
                
                If Len(MsjGritar) > 0 Then
                    Call WriteTalk("-" & MsjGritar, 2)
                End If
            End If
            
            CurrentUser.SendingType = 2
            
        Case ";" 'Globals
            If LenB(Right$(stxtbuffer, Len(stxtbuffer) - 1)) > 0 And InStr(stxtbuffer, ">") = 0 Then
                
                Dim MsjGlobal As String
                MsjGlobal = Trim(mid(stxtbuffer, InStr(stxtbuffer, ";") + 1))
                
                If Len(MsjGlobal) > 0 Then
                    Call WriteTalk(";" & MsjGlobal, 3)
                End If
            End If
            
            CurrentUser.SendingType = 7
    
        Case "\"   'Privado

            str1 = Right$(stxtbuffer, Len(stxtbuffer) - 1)
            str2 = ReadField(1, str1, 32)
        
            CurrentUser.sndPrivateTo = str2
            CurrentUser.SendingType = 3
        
            If Len(str1) - Len(str2) - 1 > 0 Then
            
                Dim Mensaje As String
                Mensaje = Right$(stxtbuffer, Len(str1) - Len(str2) - 1)

                If str1 <> "" Then
                    If Len(Trim(Mensaje)) > 0 Then
                        Call WriteWhisper(CurrentUser.sndPrivateTo, Trim(Mensaje))
                    End If
                End If
                
            End If
            
        Case Else
    
            If LenB(stxtbuffer) >= 0 Then
                Call WriteTalk(Trim(stxtbuffer), 1)
            End If
            
            CurrentUser.SendingType = 1
        
        End Select
       
    Else
        
        Call AddtoRichTextBox(General_Locale_SMG(238, 0), 0, 0, 0, 0, 0, 0, 12)
    
    End If

    stxtbuffer = vbNullString
    SendTxt.Text = vbNullString
    SendTxt.Visible = False

    Exit Sub

Enviar_SendTxt_Err:
    stxtbuffer = vbNullString
    SendTxt.Text = vbNullString
    SendTxt.Visible = False
    CurrentUser.SendingType = 1
 
    Call RegistrarError(Err.number, Err.Description, "frmMain.Enviar_SendTxt", Erl)
    Resume Next
    
End Sub


'###########################################################
'                        GUI
'###########################################################



Private Sub cmdHechizos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If CentroActual <> CentroHechizos Then Exit Sub

Select Case Index
    Case 0 'Lanzar
        cmdHechizos(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]lanzar-down")
    Case 1 'Info
        cmdHechizos(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]info-down")
    Case 2 'Subir
        cmdHechizos(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]flechaarriba-down")
    Case 3 'Bajar
        cmdHechizos(3).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]flechaabajo-down")
End Select

End Sub
Private Sub cmdHechizos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If hlst.Visible = False Then Exit Sub
If UltPosInterface = Index Then Exit Sub

If UltPosInterface <> -1 Then Call RestaurarCentroActual
UltPosInterface = Index

Select Case Index
    Case 0 'lanzar
        cmdHechizos(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]lanzar-over")
    Case 1 'info
        cmdHechizos(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]info-over")
    Case 2 'Subir
        cmdHechizos(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]flechaarriba-over")
    Case 3 'Bajar
        cmdHechizos(3).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]flechaabajo-over")
End Select

End Sub
Private Sub cmdHechizos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If CentroActual <> CentroHechizos Then Exit Sub
Call Form_MouseMove(Button, Shift, X, Y)

If hlst.ListIndex = -1 Then Exit Sub

Call Audio.PlayWave(SND_CLICK)
Dim sTemp As String

Select Case Index
    Case 0 'lanzar
     If hlst.ListIndex < 0 Then Exit Sub
        If hlst.list(hlst.ListIndex) <> Locale_GUI_Frase(269) And MainTimer.Check(TimersIndex.Work, False) Then
            If ClientTCP.DeadCheck Then Exit Sub
            Call WriteCastSpell(hlst.ListIndex + 1)
            UsaMacro = True
        End If
        
    Case 1 'info
     If hlst.ListIndex <> -1 Then
        Dim i As Byte
        For i = 1 To General_Locale_Spells(0, 7) 'maximo hechizos
           If General_Locale_Spells(i, 0) = hlst.list(hlst.ListIndex) Then
                Call AddtoRichTextBox("Nombre:" & General_Locale_Spells(i, 0) & vbCrLf & General_Locale_Spells(i, 1) & vbCrLf & "Skill requerido: " & General_Locale_Spells(i, 10) & vbCrLf & "Maná requerido: " & General_Locale_Spells(i, 8) & vbCrLf & "Energía necesaria: " & General_Locale_Spells(i, 9), 0, 0, 0, 0, 0, 0, 4)
               Exit For
           End If
        Next i
     End If
     
    Case 2 'subir
    If hlst.ListIndex = 0 Then Exit Sub
    Call WriteMoveSpell(Index, hlst.ListIndex + 1)
    
    sTemp = hlst.list(hlst.ListIndex - 1)
    hlst.list(hlst.ListIndex - 1) = hlst.list(hlst.ListIndex)
    hlst.list(hlst.ListIndex) = sTemp
    hlst.ListIndex = hlst.ListIndex - 1
    
    Case 3 'bajar
    If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
    Call WriteMoveSpell(Index, hlst.ListIndex + 1)
    sTemp = hlst.list(hlst.ListIndex + 1)
    hlst.list(hlst.ListIndex + 1) = hlst.list(hlst.ListIndex)
    hlst.list(hlst.ListIndex) = sTemp
    hlst.ListIndex = hlst.ListIndex + 1
    
End Select

End Sub

Private Sub CentroHechizosRestaurar(Index As Integer)

cmdHechizos(Index).Picture = Nothing

End Sub

Private Sub SolapasRestaurar(Index As Integer)

imgCentros(Index).Picture = Nothing
imgMiniCerra.Picture = Nothing
cmdMensaje.Picture = Nothing

End Sub

Private Sub imgCentros_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If UltPosSolapas = Index Then Exit Sub

If UltPosSolapas <> -1 Then Call RestaurarCentroActual
UltPosSolapas = Index

Select Case Index
    Case 0 'Inv
        imgCentros(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[solapas]inventario-over")
    Case 1 'Hechizos
        imgCentros(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[solapas]hechizos-over")
    Case 2 'Menu
        imgCentros(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[solapas]menu-over")
End Select

End Sub

Private Sub cmdMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0 'Grupo
        cmdMenu(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]grupo-down")
    Case 1 'Estadisticas
        cmdMenu(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]estadisticas-down")
    Case 2 'Guild
        cmdMenu(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]clanes-down")
    Case 3 'Quest
        cmdMenu(3).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]quests-down")
    Case 4 'Torneos
        cmdMenu(4).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]torneos-down")
    Case 5 'Opciones
        cmdMenu(5).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]opciones-down")
End Select

End Sub

Private Sub cmdMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If UltPosInterface = Index Then Exit Sub

If UltPosInterface <> -1 Then Call RestaurarCentroActual
UltPosInterface = Index

Select Case Index

    Case 0 'Grupo
        cmdMenu(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]grupo-over")
    Case 1 'Estadisticas
        cmdMenu(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]estadisticas-over")
    Case 2 'Guild
        cmdMenu(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]clanes-over")
    Case 3 'Quest
        cmdMenu(3).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]quests-over")
    Case 4 'Torneos
        cmdMenu(4).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]torneos-over")
    Case 5 'Opciones
        cmdMenu(5).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]opciones-over")
End Select

End Sub

Private Sub cmdMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If CentroActual <> CentroMenu Then Exit Sub
Call Form_MouseMove(Button, Shift, X, Y)

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0 'Grupo
         Call AddtoRichTextBox(General_Locale_SMG(10, 0), 0, 0, 0, 0, 0, 0, 4)
    Case 1 'Estadisticas
        
        If pausa Then Exit Sub
        
        LlegaronAtrib = False '
        LlegaronSkills = False
        LlegaronStats = False
        Call WriteRequestAtributes
        Call WriteRequestSkills
        Call WriteRequestMiniStats
            
    Case 2 'Guild
            'Abrimos los clanes
            If frmGuildLeader.Visible Then Unload frmGuildLeader
            Call WriteRequestGuildLeaderInfo
           
    Case 3 'Quest
         Call AddtoRichTextBox(General_Locale_SMG(10, 0), 0, 0, 0, 0, 0, 0, 4)
    Case 4 'Torneos
         Call AddtoRichTextBox(General_Locale_SMG(10, 0), 0, 0, 0, 0, 0, 0, 4)
    Case 5 'Opciones
        Call frmOpciones.Init
End Select

End Sub

Private Sub imgCentros_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Form_MouseMove(Button, Shift, X, Y)
Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        RenderInv = True
        Call CambiaCentro(CentroInventario)
    Case 1
        Call CambiaCentro(CentroHechizos)
    Case 2
        Call CambiaCentro(CentroMenu)
End Select

End Sub

Private Sub CentroMenuRestaurar(Index As Integer)

cmdMenu(Index).Picture = Nothing

End Sub

Private Sub RestaurarCentroActual()

Select Case CentroActual
    Case CentroHechizos
        If UltPosInterface <> -1 Then Call CentroHechizosRestaurar(UltPosInterface)
    Case CentroInventario
    Case CentroMenu
        If UltPosInterface <> -1 Then Call CentroMenuRestaurar(UltPosInterface)
End Select

If UltPosSolapas <> -1 Then Call SolapasRestaurar(UltPosSolapas)

UltPosInterface = -1
UltPosSolapas = -1

imgMiniCerra.Picture = Nothing
cmdMensaje.Picture = Nothing
lblInvInfo.Caption = vbNullString
CurrentUser.LastItem = 0

End Sub
  
 
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 
 
    If (Not SendTxt.Visible) Then
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

          If charlist(CurrentUser.UserCharIndex).EsGM Then
          
            Select Case KeyCode
            
                Case vbKeyP
                    CurrentUser.RenderGM = Not CurrentUser.RenderGM
                    
                Case vbKeyB
                    If CurrentUser.RenderGM Then
                        frmGMPanel.Show , frmMain
                    End If
                    
                Case vbKeyC
                
                    If CurrentUser.RenderGM Then
                        frmBuscar.Show , frmMain
                    End If
                    
                Case vbKeyM
                
                    If CurrentUser.RenderGM Then
                        Dim Map As String
                        Map = InputBox("Ingrese numero de mapa a viajar.")
                        If IsNumeric(Map) Then If Map > 0 And Map < 864 Then Call WriteWarpChar("YO", Map, 50, 50)
                    End If
                  
                Case vbKeyF
                
                    If CurrentUser.RenderGM Then
                        Dim i As Integer
                    
                        For i = 1 To NUMFONTS
                            Call AddtoRichTextBox("FONT: " & i & " " & Locale_GUI_Frase(245), 0, 0, 0, 0, 0, 0, i)
                        Next i
                    End If
                    
            End Select
        
          End If
        
          Select Case KeyCode

            Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                 Call WriteWork(eSkill.robar)
                                     
            Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                  Call TirarItem

            Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
                     isSeguro = Not isSeguro
                    If isSeguro = True Then
                        modoseguro.Visible = True
                        nomodoseguro.Visible = False
                    Else
                        modoseguro.Visible = False
                        nomodoseguro.Visible = True
                    End If
            
                    Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
                    Call WriteCombatModeToggle
                    IScombate = Not IScombate
                    Call Mod_General.modocombate
                    
             Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    Call WriteWork(eSkill.Ocultarse)

                  
             Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    Call WriteWork(eSkill.domar)
                
             Case CustomKeys.BindedKey(eKeyType.mKeyTakeMostrarFps)
                        FPSFLAG = Not FPSFLAG

                           
             Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                If MainTimer.Check(TimersIndex.UseItemWithU) Then
                    Call UsarItem
                End If
                           
              Case CustomKeys.BindedKey(eKeyType.mKeyROL)
                Call AddtoRichTextBox(Locale_GUI_Frase(590), 0, 0, 0, 0, 0, 0, 12)
                 
               Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                Call AgarrarItem
                     
              Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
                    Call frmMain.Client_Screenshot(frmMain.hDC, 800, 600)

                    
              Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                     
              Case CustomKeys.BindedKey(eKeyType.mkeyBloqueoMovimiento)
                  CurrentUser.AutoNavigation = Not CurrentUser.AutoNavigation
 
              Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
              Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                    
             Case vbKeyF1 To vbKeyF11
                 Call UsarMacro(KeyCode - 111)
             
              
             Case vbKeyEscape
                 Call WriteQuit
                 
          End Select

        End If
    End If
    
    Select Case KeyCode
    
          
              
             Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            
                    If Shift <> 0 Then Exit Sub
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                        If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                    Else
        
                        If Not MainTimer.Check(TimersIndex.Attack) Or CurrentUser.UserDescansar Or UserMeditar Then Exit Sub
                    End If
                    
                    Call WriteAttack
                
             Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
             
                If (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
        
                     If Not SendTxt.Visible Then
                        If Not frmCantidad.Visible Then
                            Call CompletarEnvioMensajes
                            SendTxt.Visible = True
                            SendTxt.SetFocus
                        End If
                    Else
                        Call Enviar_SendTxt
                    End If
                
                             
            
                ElseIf SendTxt.Visible Then
                    SendTxt.SetFocus
                End If
        
    End Select
End Sub
Private Sub modoseguro_Click()
    Call WriteResuscitationToggle
    isSeguro = Not isSeguro
    If isSeguro = True Then
        modoseguro.Visible = True
        nomodoseguro.Visible = False
    Else
        modoseguro.Visible = False
        nomodoseguro.Visible = True
    End If
End Sub
Private Sub nomodoseguro_Click()
    Call WriteResuscitationToggle
    isSeguro = Not isSeguro
    If isSeguro = True Then
        modoseguro.Visible = True
        nomodoseguro.Visible = False
    Else
        modoseguro.Visible = False
        nomodoseguro.Visible = True
    End If
End Sub
 
 
Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub
Private Sub MainViewPic_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

If Button = vbRightButton Then 'click derecho
    Select Case accionMousedos
    Case 1 ' "Accionar/Tomar objeto"
    Call AgarrarItem
        
    Case 2 ' "Accionar/Tomar objeto"
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                Else
                    If Not MainTimer.Check(TimersIndex.Attack) Or CurrentUser.UserDescansar Or UserMeditar Then Exit Sub
                End If
            
            Call WriteAttack
            
 Case 3    ' "Accionar/Tomar objeto"
 
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                Else
                    If Not MainTimer.Check(TimersIndex.UseItemWithU) Or CurrentUser.UserDescansar Or UserMeditar Then Exit Sub
            End If
            Call UsarItem
            
 
Case 4  ' "Accionar/Tomar objeto"
 
 End Select
 
 End If
 
 
 If Button = vbLeftButton Then
  
 Select Case accionMouseUno
 
 Case 1 ' "Accionar/Tomar objeto"
 Call AgarrarItem
 
 Case 2    ' "Accionar/Tomar objeto"
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                Else
                    If Not MainTimer.Check(TimersIndex.Attack) Or CurrentUser.UserDescansar Or UserMeditar Then Exit Sub
                End If
            Call WriteAttack
            
  Case 3 '"Accionar/Tomar objeto"
         If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                Else
                    If Not MainTimer.Check(TimersIndex.UseItemWithU) Or CurrentUser.UserDescansar Or UserMeditar Then Exit Sub
                End If
         
                Call UsarItem
            
End Select
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If

End Sub
Private Sub Minimap_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
   If X > 87 Then X = 86
   If X < 14 Then X = 15
   If Y > 90 Then Y = 89
   If Y < 11 Then Y = 12
   If Button = vbRightButton Then
      Call WriteWarpChar("YO", CurrentUser.UserMap, CByte(X - 1), CByte(Y - 1))
      Call ActualizarMiniMapa
   End If
End Sub
 



Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyReturn Then
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
        BloqueoAlCaminar = False
    End If
    
End Sub
Private Sub MainViewPic_Click()

    If Not Comerciando Then
        
        If Rendimiento = 0 Then Call ConvertCPtoTP(frmMain.MouseX, frmMain.MouseY, TX, TY)

        If charlist(CurrentUser.UserCharIndex).EsGM Then
            
            If TX = 0 And TY = 0 Then Exit Sub
            
            If MapData(TX, TY).charindex <> 0 Then
                UserFichado = MapData(charlist(MapData(TX, TY).charindex).Pos.X, charlist(MapData(TX, TY).charindex).Pos.Y).charindex
            Else
                UserFichado = 0
            End If
        End If
        
        Dim ButtonChange As Integer
        Select Case accionMousedos
        Case 2
        ButtonChange = vbLeftButton
        Case Else
        ButtonChange = vbRightButton
        End Select
    
        If MouseShift = 0 Then
     
            If MouseBoton <> ButtonChange Then 'click izquierdo
                
                Select Case CurrentUser.UsingSkill
                
                Case 0
                    Call WriteLeftClick(TX, TY)
                    
                    If MapData(TX, TY).NPCIndex <> 0 Then

                        If General_Locale_NPCs((MapData(TX, TY).NPCIndex), 1) <> "" Then
                            Call RemoveDialogsNPCArea
                            Call Char_Dialog_Create(MapData(TX, TY).charindex, General_Locale_NPCs((MapData(TX, TY).NPCIndex), 1), -1)
                        End If
                    
                    End If

                    If MapData(TX, TY).OBJInfo.OBJIndex <> 0 Then
                        If MostrarCantidad(MapData(TX, TY).OBJInfo.OBJIndex) Then
                            Call AddtoRichTextBox(General_Locale_Obj(MapData(TX, TY).OBJInfo.OBJIndex, 0) & IIf(Len(General_Locale_Obj(MapData(TX, TY).OBJInfo.OBJIndex, 1)) > 0, " - " & General_Locale_Obj(MapData(TX, TY).OBJInfo.OBJIndex, 1) & ".", "") & " (" & MapData(TX, TY).OBJInfo.Amount & ")", 0, 0, 0, 0, 0, 0, 4)
                        Else
                            Call AddtoRichTextBox(General_Locale_Obj(MapData(TX, TY).OBJInfo.OBJIndex, 0), 0, 0, 0, 0, 0, 0, 4)
                        End If
                    End If
                    
                            
                 Case proyectiles, armasarrojadizas
                 
                     If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        Call FormParser.Parse_Form(Me)
                        CurrentUser.UsingSkill = 0
                        Exit Sub
                    End If
                    
                    If Not MainTimer.Check(TimersIndex.Arrows) Then
                        Call FormParser.Parse_Form(Me)
                        CurrentUser.UsingSkill = 0
                        Exit Sub
                    End If
                 
                                     
                    Call FormParser.Parse_Form(Me)
                    Call WriteWorkLeftClick(TX, TY, CurrentUser.UsingSkill)
                    CurrentUser.UsingSkill = 0
                    
                 Case magia
                 
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                       ' Call FormParser.Parse_Form(Me)
                       ' CurrentUser.UsingSkill = 0
                        Exit Sub
                    End If
            
                   If Not MainTimer.Check(TimersIndex.CastSpell) Then
                      ' Call FormParser.Parse_Form(frmMain)
                      ' CurrentUser.UsingSkill = 0
                       Exit Sub
                   End If
               
                   Call FormParser.Parse_Form(Me)
                   Call WriteWorkLeftClick(TX, TY, CurrentUser.UsingSkill)
                   CurrentUser.UsingSkill = 0
               
                 Case pesca, robar, talar, mineria, FundirMetal, domar
                
                 
                     If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        Call FormParser.Parse_Form(Me)
                        CurrentUser.UsingSkill = 0
                        Exit Sub
                    End If
                    
                   If Not MainTimer.Check(TimersIndex.Work) Then
                           Call FormParser.Parse_Form(Me)
                           CurrentUser.UsingSkill = 0
                           Exit Sub
                       End If
                    
                    Call FormParser.Parse_Form(Me)
                    Call WriteWorkLeftClick(TX, TY, CurrentUser.UsingSkill)
                    CurrentUser.UsingSkill = 0
                
                
                Case Else
                    
                    If CurrentUser.UsingSkill = 0 Then
                            
                            Call WriteLeftClick(TX, TY)
                            
                            If MapData(TX, TY).NPCIndex <> 0 Then
                                
                                If General_Locale_NPCs((MapData(TX, TY).NPCIndex), 1) <> "" Then _
                                    Call Char_Dialog_Create(MapData(TX, TY).charindex, General_Locale_NPCs((MapData(TX, TY).NPCIndex), 1), -1)
                            End If
                            
                            If MapData(TX, TY + 1).NPCIndex <> 0 Then
                             
                                If General_Locale_NPCs((MapData(TX, TY + 1).NPCIndex), 1) <> "" Then _
                                    Call Char_Dialog_Create(MapData(TX, TY + 1).charindex, General_Locale_NPCs((MapData(TX, TY + 1).NPCIndex), 1), -1)
                            End If
                            
                            If MapData(TX, TY).OBJInfo.OBJIndex <> 0 Then
                                If MostrarCantidad(MapData(TX, TY).OBJInfo.OBJIndex) Then
                                    Call AddtoRichTextBox(General_Locale_Obj(MapData(TX, TY).OBJInfo.OBJIndex, 0) & " - " & MapData(TX, TY).OBJInfo.Amount, 204, 193, 115, False, True, False)
                                Else
                                    Call AddtoRichTextBox(General_Locale_Obj(MapData(TX, TY).OBJInfo.OBJIndex, 0), 204, 193, 115, False, True, False)
                                End If
                            End If
                    End If
                End Select
 
            Else
                If Not frmComerciar.Visible And Not frmBancoObj.Visible And Not frmCorreo.Visible Then
                    Call WriteDoubleClick(TX, TY)
                End If
            End If
        ElseIf MouseBoton = ButtonChange Then
            Call WriteDoubleClick(TX, TY)
        
        End If
    End If
    
    If charlist(CurrentUser.UserCharIndex).EsGM Then
    
        If MouseShift = vbLeftButton Then
            Call WriteWarpChar("YO", CurrentUser.UserMap, TX, TY)
        End If
        
    End If
        
End Sub
Private Sub MainViewPic_DblClick()
    
    Select Case MouseBoton
    
    Case vbRightButton
    
        If MapData(TX, TY).charindex <> 0 Then
        
            If Len(charlist(MapData(TX, TY).charindex).Nombre) <> 0 Then
                    
                    If charlist(MapData(TX, TY).charindex).EsUsuario Then
                    
                       SendTxt.Visible = True
                       SendTxt.SetFocus
                       
                       CurrentUser.sndPrivateTo = charlist(MapData(TX, TY).charindex).Nombre
                       
                       SendTxt.Text = ("\" & CurrentUser.sndPrivateTo & " ")
           
                       stxtbuffer = SendTxt.Text
                       SendTxt.SelStart = Len(SendTxt.Text)
                    
                    End If
            End If
        End If
 
    End Select
    
End Sub
 
Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub RecTxt_Change()

    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar

    If Not Application.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
         
        If picInv.Visible Then
            picInv.SetFocus
 
    End If
    End If

End Sub

Private Sub RecTxt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo RecTxt_MouseUp_Err
    
    If Button = 1 Then

        Dim strBuffer      As String

        Dim lngLength      As Long

        Dim intCurrentLine As Integer
    
        intCurrentLine = RecTxt.GetLineFromChar(RecTxt.SelStart)
        'get line length
        lngLength = SendMessage(RecTxt.hwnd, EM_LINELENGTH, intCurrentLine, 0)
        'resize buffer
        strBuffer = Space(lngLength)
        'get line text
        Call SendMessage(RecTxt.hwnd, EM_GETLINE, intCurrentLine, ByVal strBuffer)

        Dim partea       As String

        Dim destinatario As String
    
        destinatario = SuperMid(strBuffer, "[", "]", False)

        If destinatario <> "A" Then

            destinatario = Replace(destinatario, " ", "+")

            CurrentUser.sndPrivateTo = destinatario
            SendTxt.Text = ("\" & CurrentUser.sndPrivateTo & " ")

            stxtbuffer = SendTxt.Text
            SendTxt.SelStart = Len(SendTxt.Text)

            If SendTxt.Visible = False Then
             '  Call WriteEscribiendo

            End If

            SendTxt.Visible = True
            SendTxt.SetFocus

        End If

    End If

    
    Exit Sub

RecTxt_MouseUp_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMain.RecTxt_MouseUp", Erl)
    Resume Next
    
End Sub

Public Sub Client_Screenshot(ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long)

On Error GoTo errorhandler

Dim i As Long
Dim Index As Long
i = 1

Set m_Jpeg = New clsJpeg

'80 Quality
m_Jpeg.Quality = 100

'Sample the cImage by hDC
m_Jpeg.SampleHDC hDC, Width, Height

m_FileName = App.Path & "\Fotos\CovAO_Foto"

If Dir(App.Path & "\Fotos", vbDirectory) = vbNullString Then
    MkDir (App.Path & "\Fotos")
End If

Do While Dir(m_FileName & Trim(str(i)) & ".jpg") <> vbNullString
    i = i + 1
    DoEvents
Loop

Index = i
 
'Save the JPG file
m_Jpeg.SaveFile m_FileName & Trim(str(Index)) & ".jpg"
'"Screenshot grabada correctamente como
 
 AddtoRichTextBox "Screenshot grabada correctamente como " & " " & m_FileName & Trim(str(Index)) & ".jpg", 65, 190, 156, False, False, True
         
Set m_Jpeg = Nothing

Exit Sub

errorhandler:
  AddtoRichTextBox frmMain.RecTxt, "Error al grabar el screenshot. Por favor intente nuevamente.", 65, 190, 156, False, False, True

End Sub
Public Sub RenderMacro(ByRef Pic As PictureBox, ByVal GrhIndex As Long)
    Dim SR As Rect
    Dim DR As Rect
  
    With GrhData(GrhIndex)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .pixelWidth
        SR.bottom = SR.Top + .pixelHeight
    End With
  
    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.bottom = 32
  
    Call DrawGrhtoHdc(Pic.hDC, GrhIndex, SR.Left, SR.Top)
    Pic.Refresh
End Sub
Private Sub nuevocorreo_Click()
 Call AddtoRichTextBox(Locale_GUI_Frase(591), 0, 0, 0, 0, 0, 0, 12)
End Sub

Private Sub shape2_click(Index As Integer)
 If Valor = 0 Then
Label1(Index).Visible = True
 
 Valor = 1
Else '
'  If Valor = 1 Then
 Label1(Index).Visible = False
 Valor = 0
 End If
End Sub
 

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
Private Sub Socket1_Connect()

    Debug.Print "Open Socket1.Connect: " & Time
    
     Socket1.NoDelay = True
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
    
    'Limpio el codigo
    Security.Redundance = 13
    
    Select Case EstadoLogin
        
    Case E_MODO.CrearNuevoPj, E_MODO.Normal, E_MODO.ConectarPersonaje, E_MODO.BorrarPersonaje
        Call Login
        
    Case E_MODO.CrearNuevaCuenta
        
        If frmCrearCuenta.Visible = True Then
            Call Login
            
        Else
            Call frmCrearCuenta.InitCuenta
        
        End If
        
    Case E_MODO.Dados
    
        frmCrearPersonaje.Show
        
    Case E_MODO.RecuperarCuenta
    
        If frmRecuperarCuenta.Visible = True Then
            Call Login
        Else
            frmRecuperarCuenta.Show vbModal, frmConnect
        End If
        
    Case E_MODO.CambiarContraseña
    
        If frmCambiarContraseña.Visible = True Then
            Call Login
        Else
            frmCambiarContraseña.Show vbModal, frmCharList
        End If
        
    End Select

    End Sub
    

Public Sub Socket_Error_Close_Event(Optional ByVal Description As String, Optional ByVal number As Long)

    Dim bFlag As Boolean
        
    frmMain.Socket1.Cleanup
    
    Connected = False
    RenderInv = False

    
    If frmMensaje.Visible And frmConnect.Visible = False Then
        frmMensaje.Visible = False
        bFlag = True
    End If

    If CurrentUser.Logged Then
    
        Call ResetCharDisconnect
        
        frmConnect.Visible = True
        Me.Visible = False
        frmCharList.Visible = False
        frmIniciando.Visible = False
        frmCrearPersonaje.Visible = False
        
        Call FormParser.Parse_Form(frmConnect)
        
        CurrentUser.LogeoAlgunaVez = False
        
    Else
    
        frmConnect.Visible = True
        Me.Visible = False
        frmCharList.Visible = False
        frmIniciando.Visible = False
        frmCrearPersonaje.Visible = False
        
        Call FormParser.Parse_Form(frmConnect)

    End If
    
    If bFlag Then
    
        If frmConnect.Visible Then
            frmMensaje.Show vbModal, frmConnect
        ElseIf frmCharList.Visible Then
            frmMensaje.Show vbModal, frmCharList
        ElseIf frmCrearCuenta.Visible Then
            frmMensaje.Show vbModal, frmCrearCuenta
        ElseIf frmRecuperarCuenta.Visible Then
            frmMensaje.Show vbModal, frmRecuperarCuenta
        ElseIf frmCambiarContraseña.Visible Then
            frmMensaje.Show vbModal, frmCambiarContraseña
        End If
    Else
    
        If Not frmMensaje.Visible Then
        
            If LenB(Description) > 0 Then
                Select Case number
                Case 24036
                    frmMensaje.msg.Caption = Locale_Error(43)
                    frmMensaje.Show , frmConnect
                
                Case 24061, 24053
                    frmMensaje.msg.Caption = Locale_Error(27)
                    frmMensaje.Show , frmConnect
                    
                Case Else
                    frmMensaje.msg.Caption = (Locale_GUI_Frase(345) & " (" & Description & " - " & number & ")")
                    frmMensaje.Show vbModal, frmConnect
                    
                End Select
                'Call MsgBox(Locale_GUI_Frase(345) & " (" & Description & " - " & Number & ")", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error al conectar")
                 
                 
            
            End If
        End If
        
    End If
    
    Debug.Print "Finish Socket_Error_Close_Event " & Time
    
    
End Sub

Private Sub Socket1_Disconnect()

    Debug.Print "Cerrando la conexion Socket1.Disconnect" & Time
    Call Socket_Error_Close_Event
    
End Sub
Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Debug.Print "Open Socket1_LastError: " & Time
    
        Debug.Print "Estamos en socket1_last error mientras estamos conectados, abrimos socket1_disconnect"
                        
        If Connected = True Then
            CurrentUser.LogeoAlgunaVez = True
        End If
        
        Socket1.Disconnect
 
        Call Socket_Error_Close_Event(ErrorString, ErrorCode)
      
    Err.Clear
End Sub
 
Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)

    Dim RD     As String

    Dim data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
 
    'Put data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
    
End Sub

Private Sub Socket1_Timeout(status As Integer, Response As Integer)
    
    On Error GoTo Socket1_Timeout_Err
    
    Debug.Print "Open Socket1.timeOut" & Time

    
    Exit Sub

Socket1_Timeout_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMain.Socket1_Timeout", Erl)
    Resume Next
    
End Sub
 
Public Sub ResetCharDisconnect()
'Estado = 1 es cuando se cierra desde el panel cuenta
'Estado = 0 es cuando se reinicia el socket por errores
        
    Debug.Print "Open ResetCharDisconnect " & Time
    
    Dim i As Long
    
    'Reset global vars
    EngineRun = False
    isSeguro = False
    IScombate = False
    CurrentUser.RenderGM = False
    CurrentUser.TiempoSalida = False
    UserParalizado = False
    pausa = False
    UserCiego = False
    UserMeditar = False
    UserNavegando = False
    CurrentUser.Montando = False
    CurrentUser.UserDescansar = False
    bRain = False
    bFogata = False
    SkillPoints = 0
    Comerciando = False
    UserCiego = False
    UserEstupido = False
    CurrentUser.LastItem = 0
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    SkillPoints = 0
    Alocados = 0
    
    scroll_pixels_per_frame = scroll_pixels_per_frameBackUp
    For i = 1 To NUMSKILLS
        CurrentUser.UserSkills(i) = 0
    Next i
    
    For i = 1 To NUMATRIBUTOS
        CurrentUser.UserAtributos(i) = 0
    Next i
    
    For i = 1 To MAX_INVENTORY_SLOTS
        Call Inventario.SetItem(i, 0, 0, 0, 0, 0, True)
    Next i
    
    'Reset some char variables...
    For i = 1 To LastChar
        charlist(i).Invisible = False
        Call ResetCharInfo(i)
        Call Char_Dialog_Remove(i)
    Next i
    
    'Call RefreshAllChars
    
    'Stop audio
    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone
    
End Sub
Public Function EquiparObjeto(ByVal Slot As Byte)

   
If Not charlist(CurrentUser.UserCharIndex).EsGM Then

    Dim Nivel As Integer
    Nivel = General_Locale_Obj(Inventario.OBJIndex(Inventario.SelectedItem), 14)
    If Nivel > 0 Then
        If CurrentUser.UserLvl < Nivel Then
            Call AddtoRichTextBox(Locale_Parse_ServidorMensaje(268, Nivel), 0, 0, 0, 0, 0, 0, 12)
            Exit Function
        End If
    End If
    
End If
    
Call WriteEquipItem(Slot)

End Function
