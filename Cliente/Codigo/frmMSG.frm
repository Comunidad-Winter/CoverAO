VERSION 5.00
Begin VB.Form frmMSG 
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   45
   ClientWidth     =   7245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMSG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "$2"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton imgCerrar 
      Caption         =   "$2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
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
      Height          =   1785
      Left            =   300
      TabIndex        =   0
      Top             =   615
      Width           =   1845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta de:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmMSG"
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

 
 



Private Sub Command1_Click()

Call WriteCleanSOS
Unload Me
End Sub

Private Sub Form_Deactivate()
 
    Me.Visible = False
    List1.Clear

End Sub

Private Sub Form_Load()
    List1.Clear
    
'    Me.Picture = LoadPicture(App.path & "\graficos\VentanaShowSos.jpg")
    Call FormParser.Parse_Form(Me)

 
End Sub

Private Sub imgCerrar_Click()
    Me.Visible = False
    List1.Clear

End Sub

Private Sub list1_Click()

    Dim ind As Integer

    ind = val(ReadField(2, List1.list(List1.ListIndex), Asc("-")))

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopupMenu menU_usuario

    End If

End Sub

Private Sub mnuBorrar_Click()

    If List1.ListIndex < 0 Then Exit Sub

    'Pablo (ToxicWaste)
    Dim Aux As String

    Aux = mid$(ReadField(1, List1.list(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.list(List1.ListIndex), Asc("-"))))
    Call WriteSOSRemove(Aux)
    '/Pablo (ToxicWaste)
    'Call WriteSOSRemove(List1.List(List1.listIndex))
    
    List1.RemoveItem List1.ListIndex

End Sub

Private Sub mnuIR_Click()

    'Pablo (ToxicWaste)
    Dim Aux As String

    Aux = mid$(ReadField(1, List1.list(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.list(List1.ListIndex), Asc("-"))))
    Call WriteGoToChar(Aux)
    '/Pablo (ToxicWaste)
    'Call WriteGoToChar(ReadField(1, List1.List(List1.listIndex), Asc("-")))
    
End Sub

Private Sub mnutraer_Click()

    'Pablo (ToxicWaste)
    Dim Aux As String

    Aux = mid$(ReadField(1, List1.list(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.list(List1.ListIndex), Asc("-"))))
    Call WriteSummonChar(Aux)

    'Pablo (ToxicWaste)
    'Call WriteSummonChar(ReadField(1, List1.List(List1.listIndex), Asc("-")))
End Sub

