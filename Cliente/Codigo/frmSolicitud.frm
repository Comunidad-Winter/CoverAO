VERSION 5.00
Begin VB.Form frmGuildSol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4350
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
   Icon            =   "frmSolicitud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   236
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   290
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   120
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton imgEnviar 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   2280
      MouseIcon       =   "frmSolicitud.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton imgCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   240
      MouseIcon       =   "frmSolicitud.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text11 
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
      Height          =   1035
      Left            =   5220
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmSolicitud.frx":02B0
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
   Begin VB.Image imgEnviar1 
      Height          =   525
      Left            =   8280
      Tag             =   "1"
      Top             =   2880
      Width           =   945
   End
   Begin VB.Image imgCerrar1 
      Height          =   525
      Left            =   5160
      Tag             =   "1"
      Top             =   2880
      Width           =   945
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

 
Dim CName             As String

Public Sub RecieveSolicitud(ByVal GuildName As String)

    CName = GuildName

End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
 
   ' Me.Picture = LoadPicture(App.path & "\graficos\VentanaIngreso.jpg")
    
 
End Sub
 
Private Sub imgCerrar_Click()
    Unload Me

End Sub

Private Sub imgEnviar_Click()
    Call WriteGuildRequestMembership(CName, Replace(Replace(Text1.Text, ",", ";"), vbCrLf, "�"))

    Unload Me

End Sub

