VERSION 5.00
Begin VB.Form frmSkins 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   Icon            =   "frmSkins.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   360
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   240
      Width           =   1140
   End
   Begin VB.Image Image15 
      Height          =   195
      Left            =   1500
      Top             =   1960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image14 
      Height          =   195
      Left            =   1500
      Top             =   2680
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image13 
      Height          =   195
      Left            =   1500
      Top             =   3050
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image12 
      Height          =   195
      Left            =   1500
      Top             =   3400
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   165
      Left            =   1005
      TabIndex        =   0
      Top             =   1650
      Width           =   45
   End
   Begin VB.Image Image11 
      Height          =   195
      Left            =   1500
      Top             =   2320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image10 
      Height          =   255
      Left            =   960
      Top             =   3360
      Width           =   255
   End
   Begin VB.Image Image9 
      Height          =   255
      Left            =   960
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   960
      Top             =   2640
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   960
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   960
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   0
      Top             =   3360
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   0
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   0
      Top             =   2640
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   0
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Top             =   1920
      Width           =   255
   End
End
Attribute VB_Name = "frmSkins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Form_Load()
 
    Call FormParser.Parse_Form(Me)
 
    Make_Transparent_Form Me.hwnd, 180

    Label3.Caption = CurrentUser.Creditos
     
 End Sub
 
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
      Unload Me
End If

End Sub
Private Sub Image1_Click()
Call WriteEquiparSkin("Cascos", 0)
End Sub

Private Sub Image10_Click()
Call WriteEquiparSkin("Objetos", 1)
End Sub

Private Sub Image2_Click()
Call WriteEquiparSkin(1, 0) 'Armaduras / Anterior
End Sub
Private Sub Image7_Click()
Call WriteEquiparSkin(1, 1)  'Armaduras / Siguiente
End Sub

Private Sub Image3_Click()
Call WriteEquiparSkin("Escudos", 0)
End Sub

Private Sub Image4_Click()
Call WriteEquiparSkin("Armas", 0)
End Sub

Private Sub Image5_Click()
Call WriteEquiparSkin("Objetos", 0)

End Sub

Private Sub Image6_Click()
Call WriteEquiparSkin("Cascos", 1)
End Sub

 

Private Sub Image8_Click()
Call WriteEquiparSkin("Escudos", 1)
End Sub

Private Sub Image9_Click()
Call WriteEquiparSkin("Armas", 1)
End Sub
