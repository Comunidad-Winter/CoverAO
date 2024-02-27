VERSION 5.00
Begin VB.Form frmSpawnList 
   BorderStyle     =   0  'None
   Caption         =   "$424"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSpawnList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "$97"
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.ListBox lstCriaturas 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   120
      TabIndex        =   2
      Top             =   330
      Width           =   2355
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "$1"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2970
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "$2"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   0
      Top             =   2970
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "$98"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmSpawnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

 

Private Sub cmdButton_Click(Index As Integer)
 
 Select Case Index
   
   Case 0
      Unload Me
   
   Case 1
   If MsgBox(Locale_GUI_Frase(340), vbYesNo) = vbNo Then Exit Sub
   
     Call WriteSpawnCreature(lstCriaturas.ListIndex + 1)
   
 End Select
 
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
 
  '  Me.Picture = LoadPicture(App.path & "\graficos\VentanaInvocar.jpg")
 '
End Sub

