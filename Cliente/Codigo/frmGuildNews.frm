VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GuildNews"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4830
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmGuildNews.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton imgAceptar 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmGuildNews.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5880
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "GuildNews"
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4575
      Begin VB.TextBox news 
         ForeColor       =   &H00000000&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clanes con los que estamos en guerra"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   4575
      Begin VB.ListBox txtClanesGuerra 
         Height          =   1035
         ItemData        =   "frmGuildNews.frx":015E
         Left            =   120
         List            =   "frmGuildNews.frx":0165
         TabIndex        =   6
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes aliados"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   4575
      Begin VB.ListBox txtClanesAliados 
         Height          =   1035
         ItemData        =   "frmGuildNews.frx":017A
         Left            =   120
         List            =   "frmGuildNews.frx":0181
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.TextBox txtClanesAliados1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   1020
      Left            =   6075
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4800
      Width           =   4275
   End
   Begin VB.TextBox txtClanesGuerra1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   1020
      Left            =   6075
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3240
      Width           =   4275
   End
   Begin VB.TextBox news1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   2100
      Left            =   6075
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   585
      Width           =   4275
   End
   Begin VB.Image imgAceptar1 
      Height          =   375
      Left            =   6075
      Tag             =   "1"
      Top             =   6000
      Width           =   4350
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

 
Private Sub aliados_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)
    

End Sub

Private Sub Form_Load()
 
    'Me.Picture = LoadPicture(App.path & "\graficos\VentanaGuildNews.jpg")
    
 
End Sub
 
Private Sub imgAceptar_Click()

    On Error Resume Next

    Unload Me
    frmMain.SetFocus

End Sub

Private Sub imgAceptar_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 y As Single)
 
End Sub

