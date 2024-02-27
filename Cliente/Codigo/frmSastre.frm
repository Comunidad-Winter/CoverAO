VERSION 5.00
Begin VB.Form frmSastre 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$422"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4920
   Icon            =   "frmSastre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "$2"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmSastre.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3000
      Width           =   1710
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
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmSastre.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3000
      Width           =   1710
   End
   Begin VB.ListBox lstRopas 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4665
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "1"
      Top             =   2640
      Width           =   4665
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "$22"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   4665
   End
End
Attribute VB_Name = "frmSastre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
    On Error Resume Next

    Call WriteCraftSastre(ObjSastre(lstRopas.ListIndex + 1), val(txtCantidad.Text))

    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)

End Sub

Private Sub txtCantidad_Change()
    If val(txtCantidad.Text) < 0 Then
        txtCantidad.Text = 1
    End If
    
    If val(txtCantidad.Text) > 1000 Then
        txtCantidad.Text = 1
    End If
End Sub
