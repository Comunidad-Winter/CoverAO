VERSION 5.00
Begin VB.Form frmGuildURL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oficial Web Site"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   6225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGuildURL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUrl 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5895
   End
   Begin VB.CommandButton imgAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmGuildURL.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   720
      Width           =   5895
   End
   Begin VB.TextBox txtUrl1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   210
      TabIndex        =   0
      Top             =   1920
      Width           =   5805
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese la direccion del site:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image imgAceptar1 
      Height          =   255
      Left            =   165
      Tag             =   "1"
      Top             =   2280
      Width           =   5880
   End
End
Attribute VB_Name = "frmGuildURL"
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
 

Private Sub Form_Load()
 
   ' Me.Picture = LoadPicture(App.path & "\graficos\VentanaUrlClan.jpg")
    
 
End Sub
 
Private Sub imgAceptar_Click()

    If txtUrl.Text <> "" Then Call WriteGuildNewWebsite(txtUrl.Text)
    
    Unload Me

End Sub

Private Sub txtUrl_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)
   

End Sub
