VERSION 5.00
Begin VB.Form frmPres 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "CoverAO 1.0"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPres.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrpres 
      Interval        =   6000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
' soporte@coverao.com.ar
'   - Relase Number 1
'*****************************************************************

Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = Form_Caption
Call FormParser.Parse_Form(Me)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
FinPres = True
End Sub

Private Sub tmrPres_Timer()
FinPres = True
End Sub



 


