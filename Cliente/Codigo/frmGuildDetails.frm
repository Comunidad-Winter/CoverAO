VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalles del Clan"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6735
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
   Icon            =   "frmGuildDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmDesc 
      Caption         =   "$39"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Codex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   6495
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   1560
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   6
         Top             =   2640
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   5
         Left            =   360
         TabIndex        =   5
         Top             =   3000
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   6
         Left            =   360
         TabIndex        =   4
         Top             =   3360
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   7
         Left            =   360
         TabIndex        =   3
         Top             =   3720
         Width           =   5655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"frmGuildDetails.frx":000C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.CommandButton imgSalir 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildDetails.frx":0145
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton imgConfirmar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5160
      MouseIcon       =   "frmGuildDetails.frx":0297
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6360
      Width           =   1455
   End
End
Attribute VB_Name = "frmGuildDetails"
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

 Private Const MAX_DESC_LENGTH  As Integer = 520

Private Const MAX_CODEX_LENGTH As Integer = 100

Private Sub Form_Load()
  
 '   Me.Picture = LoadPicture(App.path & "\graficos\VentanaCodex.jpg")
    
 
End Sub
 
Private Sub imgConfirmar_Click()

    Dim fdesc   As String

    Dim Codex() As String

    Dim k       As Byte

    Dim Cont    As Byte

    fdesc = Replace(txtDesc, vbCrLf, "�", , , vbBinaryCompare)

    Cont = 0

    For k = 0 To txtCodex1.UBound

        If LenB(txtCodex1(k).Text) <> 0 Then Cont = Cont + 1
    Next k
    
    If Cont < 4 Then
        MsgBox "Debes definir al menos cuatro mandamientos."
        Exit Sub

    End If
                
    ReDim Codex(txtCodex1.UBound) As String

    For k = 0 To txtCodex1.UBound
        Codex(k) = txtCodex1(k)
    Next k

    If CreandoClan Then
        Call WriteCreateNewGuild(fdesc, ClanName, Site, Codex)
    Else
        Call WriteClanCodexUpdate(fdesc, Codex)

    End If

    CreandoClan = False
    Unload Me

End Sub

Private Sub imgSalir_Click()
    Unload Me

End Sub

Private Sub txtCodex1_Change(Index As Integer)

    If Len(txtCodex1.item(Index).Text) > MAX_CODEX_LENGTH Then txtCodex1.item(Index).Text = Left$(txtCodex1.item(Index).Text, MAX_CODEX_LENGTH)

End Sub

Private Sub txtCodex1_MouseMove(Index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
 
End Sub

Private Sub txtDesc_Change()

    If Len(txtDesc.Text) > MAX_DESC_LENGTH Then txtDesc.Text = Left$(txtDesc.Text, MAX_DESC_LENGTH)

End Sub
