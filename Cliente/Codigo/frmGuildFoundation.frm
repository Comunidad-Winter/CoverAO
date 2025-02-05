VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creaci�n de un Clan"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4095
   ClipControls    =   0   'False
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
   Icon            =   "frmGuildFoundation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Informaci�n b�sica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtClanName 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del clan:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   $"frmGuildFoundation.frx":000C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   3495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Web site del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3855
      Begin VB.TextBox txtWeb 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.CommandButton imgSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmGuildFoundation.frx":00D6
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton imgCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildFoundation.frx":0228
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
End
Attribute VB_Name = "frmGuildFoundation"
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
 
Private Sub Form_Deactivate()
    Me.SetFocus

End Sub

Private Sub Form_Load()
 
    If Len(txtClanName.Text) <= 22 Then
        If Not AsciiValidos(txtClanName) Then
            MsgBox "Nombre invalido."
            Exit Sub

        End If

    Else
        MsgBox "Nombre demasiado extenso."
        Exit Sub

    End If

End Sub
 
Private Sub imgCancelar_Click()
    Unload Me

End Sub

Private Sub imgSiguiente_Click()
    ClanName = txtClanName.Text
    Site = txtWeb.Text
    Unload Me
    frmGuildDetails.Show , frmMain

End Sub

