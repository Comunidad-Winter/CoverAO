VERSION 5.00
Begin VB.Form frmHlp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "$602"
   ClientHeight    =   6900
   ClientLeft      =   2355
   ClientTop       =   1845
   ClientWidth     =   5640
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHlp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4762.502
   ScaleMode       =   0  'User
   ScaleWidth      =   5296.251
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame Ayuda 
      Caption         =   "$602"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.Frame Frame5 
         Caption         =   "$601"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   5415
         Begin VB.TextBox txtComandos 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3195
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   2160
            Width           =   5175
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            Index           =   0
            X1              =   480
            X2              =   5050
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "$619"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   25
            Top             =   1800
            Width           =   3735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "$620"
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   5175
         End
         Begin VB.Label Label1 
            Caption         =   "$621"
            Height          =   795
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   5085
         End
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$612"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3720
         TabIndex        =   20
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "$623"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   5175
         Begin VB.TextBox txtMsg 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1395
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$613"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   3720
         TabIndex        =   13
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$614"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   3720
         TabIndex        =   12
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$611"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   3720
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$610"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1920
         TabIndex        =   10
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$609"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1920
         TabIndex        =   9
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$608"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1920
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$607"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$606"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$605"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$604"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "$603"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblHlp 
         Caption         =   "$624"
         Height          =   4995
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5085
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "$618"
      Height          =   195
      Left            =   2040
      TabIndex        =   19
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "$25"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   6480
      Width           =   5415
   End
   Begin VB.CommandButton cmdNPage 
      Caption         =   "$616"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   17
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdGM 
      Caption         =   "$617"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdBPage 
      Caption         =   "$615"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmHlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EnQuePagina As Integer

Private Sub cmdBoton_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To cmdBoton.UBound
        cmdBoton(i).Visible = False
    Next i
    
    Me.Frame1.Visible = False
    cmdNPage.Visible = False
    
    Call CambioPagina(Index)
End Sub

Private Sub cmdBPage_Click()

    Dim i As Integer
    
    For i = 1 To cmdBoton.UBound
        cmdBoton(i).Visible = True
    Next i
    
    Me.Frame1.Visible = True
    cmdNPage.Visible = True
    
    Call CambioPagina(0)
    
    Frame5.Visible = False
    
End Sub

Private Sub cmdGM_Click()
frmHlp.Visible = False
frmGMAyuda.Show vbModeless, frmMain
End Sub


Private Sub CambioPagina(ByVal EnQuePagina As Integer)

Select Case EnQuePagina

Case 0
    Ayuda.Caption = Locale_GUI_Frase(602)
    lblHlp.Caption = Locale_GUI_Frase(622)
    cmdBPage.Visible = False
Case 1
    Ayuda.Caption = Locale_GUI_Frase(603)
    lblHlp.Caption = Locale_GUI_Frase(625)
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 2
    Ayuda.Caption = Locale_GUI_Frase(604)
    lblHlp.Caption = Locale_GUI_Frase(626)
    cmdBPage.Visible = True
Case 3
    Ayuda.Caption = Locale_GUI_Frase(605)
    lblHlp.Caption = Locale_GUI_Frase(627)
    cmdBPage.Visible = True
Case 4
    Ayuda.Caption = Locale_GUI_Frase(606)
    lblHlp.Caption = Locale_GUI_Frase(628)
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 5
    Ayuda.Caption = Locale_GUI_Frase(607)
    lblHlp.Caption = Locale_GUI_Frase(629)
    cmdBPage.Visible = True
Case 6
    Ayuda.Caption = Locale_GUI_Frase(608)
    lblHlp.Caption = Locale_GUI_Frase(630)
    cmdBPage.Visible = True
Case 7
    Ayuda.Caption = Locale_GUI_Frase(609)
    lblHlp.Caption = Locale_GUI_Frase(631)
    cmdBPage.Visible = True
Case 8
    Ayuda.Caption = Locale_GUI_Frase(610)
    lblHlp.Caption = Locale_GUI_Frase(632)
    Frame5.Visible = True
    cmdBPage.Visible = True
    
Case 9
    Ayuda.Caption = Locale_GUI_Frase(611)
    lblHlp.Caption = Locale_GUI_Frase(633)
    cmdBPage.Visible = True
Case 10
    Ayuda.Caption = Locale_GUI_Frase(612)
    lblHlp.Caption = Locale_GUI_Frase(634)
    cmdBPage.Visible = True
Case 11
    Ayuda.Caption = Locale_GUI_Frase(613)
    lblHlp.Caption = Locale_GUI_Frase(635)
    cmdBPage.Visible = True
Case 12
    Ayuda.Caption = Locale_GUI_Frase(614)
    lblHlp.Caption = Locale_GUI_Frase(636)
    cmdBPage.Visible = True
End Select
End Sub

Private Sub Form_Load()
Call Locale_Soporte_Frase(0, 1)
Call FormParser.Parse_Form(Me)
End Sub
