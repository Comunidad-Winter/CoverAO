VERSION 5.00
Begin VB.Form frmGuildMember 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración del Clan - Miembro"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Icon            =   "frmGuildMember.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstClanes 
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
      Height          =   2790
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   2610
   End
   Begin VB.ListBox lstMiembros 
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
      Height          =   2010
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   2610
   End
   Begin VB.CommandButton imgDetalles 
      Caption         =   "Detalles"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton imgCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton cmdAbandonar 
      Caption         =   "Abandonar Clan"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtSearch 
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
      Height          =   225
      Left            =   3000
      TabIndex        =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Label lblCantMiembros 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4980
      TabIndex        =   9
      Top             =   2415
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Lista de Clanes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Lista de Miembros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad de Miembros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3120
      TabIndex        =   6
      Top             =   2400
      Width           =   1950
   End
   Begin VB.Line Line1 
      X1              =   192
      X2              =   192
      Y1              =   8
      Y2              =   232
   End
   Begin VB.Image imgCerrar1 
      Height          =   495
      Left            =   3000
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Image imgNoticias 
      Height          =   495
      Left            =   150
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Image imgDetalles1 
      Height          =   375
      Left            =   150
      Top             =   4200
      Width           =   2655
   End
End
Attribute VB_Name = "frmGuildMember"
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
 
Private Sub cmdAbandonar_Click()
'Salimos del clan :P
'Call Clienttcp.ParseUserCommand("/salirclan")
Unload Me
End Sub

Private Sub Form_Load()

   
     Me.Caption = "Administración del Clan - Miembro: " & Cuenta.UserName
End Sub
 
Private Sub imgCerrar_Click()
    Unload Me
frmMain.SetFocus
End Sub

Private Sub imgDetalles_Click()

    If lstClanes.ListIndex = -1 Then Exit Sub
    
    frmGuildBrief.EsLeader = False

    Call WriteGuildRequestDetails(lstClanes.list(lstClanes.ListIndex))

End Sub

Private Sub imgNoticias_Click()
    Call WriteShowGuildNews

End Sub

Private Sub txtSearch_Change()
    Call FiltrarListaClanes(txtSearch.Text)

End Sub

Private Sub txtSearch_GotFocus()

    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)

    End With

End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)

    Dim lIndex As Long
    
    If UBound(GuildNames) <> 0 Then

        With lstClanes
            'Limpio la lista
            .Clear
            
            .Visible = False
            
            ' Recorro los arrays
            For lIndex = 0 To UBound(GuildNames)

                ' Si coincide con los patrones
                If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                    ' Lo agrego a la lista
                    .AddItem GuildNames(lIndex)

                End If

            Next lIndex
            
            .Visible = True

        End With

    End If

End Sub

