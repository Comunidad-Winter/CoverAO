VERSION 5.00
Begin VB.Form frmHerrero 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$429"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmHerrero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstEscudos 
      Height          =   2205
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   5775
   End
   Begin VB.ListBox lstArmas 
      Height          =   2205
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   5775
   End
   Begin VB.ListBox lstCascos 
      Height          =   2205
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   5775
   End
   Begin VB.CommandButton cmdCascos 
      Caption         =   "$61"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdEscudos 
      Caption         =   "$60"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Text            =   "1"
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "$2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "$1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton cmdArmaduras 
      Caption         =   "$59"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdArmas 
      Caption         =   "$58"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox lstArmaduras 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "$22"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdEscudos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdEscudos.FontUnderline = True
 
 cmdCascos.FontUnderline = False
 cmdArmaduras.FontUnderline = False
 cmdArmas.FontUnderline = False
End Sub

Private Sub cmdArmaduras_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdArmaduras.FontUnderline = True
 
 cmdEscudos.FontUnderline = False
 cmdCascos.FontUnderline = False
 cmdArmas.FontUnderline = False
End Sub

Private Sub cmdArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdArmas.FontUnderline = True
 
 cmdEscudos.FontUnderline = False
 cmdCascos.FontUnderline = False
 cmdArmaduras.FontUnderline = False
End Sub

Private Sub cmdCascos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdCascos.FontUnderline = True
 
 cmdEscudos.FontUnderline = False
 cmdArmaduras.FontUnderline = False
 cmdArmas.FontUnderline = False
End Sub

Private Sub cmdArmas_Click()
    lstArmaduras.Visible = False
    lstEscudos.Visible = False
    lstCascos.Visible = False
    lstArmas.Visible = True
End Sub

Private Sub cmdCascos_Click()
    lstArmaduras.Visible = False
    lstEscudos.Visible = False
    lstArmas.Visible = False
    lstCascos.Visible = True
End Sub

Private Sub cmdEscudos_Click()
    lstArmaduras.Visible = False
    lstArmas.Visible = False
    lstCascos.Visible = False
    lstEscudos.Visible = True
End Sub

Private Sub cmdArmaduras_Click()
    lstArmaduras.Visible = True
    lstEscudos.Visible = False
    lstArmas.Visible = False
    lstCascos.Visible = False
End Sub

Private Sub Command3_Click()
On Error Resume Next

    Command3.FontUnderline = True
    
    If lstArmas.Visible Then
        Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.ListIndex + 1), val(txtCantidad.Text))
    ElseIf lstArmaduras.Visible Then
        Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmaduras.ListIndex + 1), val(txtCantidad.Text))
    ElseIf lstCascos.Visible Then
        Call WriteCraftBlacksmith(CascosHerrero(lstCascos.ListIndex + 1), val(txtCantidad.Text))
    ElseIf lstEscudos.Visible Then
        Call WriteCraftBlacksmith(EscudosHerrero(lstEscudos.ListIndex + 1), val(txtCantidad.Text))
    End If

    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lstArmaduras.Visible = True
    lstArmas.Visible = False
    lstCascos.Visible = False
    lstEscudos.Visible = False
    Call FormParser.Parse_Form(Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdEscudos.FontUnderline = False
 cmdCascos.FontUnderline = False
 cmdArmaduras.FontUnderline = False
 cmdArmas.FontUnderline = False
 Command3.FontUnderline = False
End Sub
