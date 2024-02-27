VERSION 5.00
Begin VB.Form frmShop 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   FillColor       =   &H00877365&
   Icon            =   "frmShop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Text            =   "1"
      Top             =   6960
      Width           =   510
   End
   Begin VB.ListBox lstInv 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   3710
      TabIndex        =   8
      Top             =   2580
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3960
      Left            =   750
      TabIndex        =   1
      Top             =   2580
      Width           =   2500
   End
   Begin VB.PictureBox PicItem 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   890
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1610
      Width           =   495
   End
   Begin VB.Image Image3 
      Enabled         =   0   'False
      Height          =   420
      Left            =   3850
      Tag             =   "1"
      Top             =   6890
      Width           =   195
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   420
      Left            =   2950
      Tag             =   "1"
      Top             =   6890
      Width           =   195
   End
   Begin VB.Image imgComprar 
      Height          =   450
      Left            =   580
      MouseIcon       =   "frmShop.frx":000C
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6850
      Width           =   2175
   End
   Begin VB.Image imgVender 
      Height          =   450
      Left            =   4230
      MouseIcon       =   "frmShop.frx":015E
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6850
      Width           =   2175
   End
   Begin VB.Label lblDefensa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   1
      Left            =   1575
      TabIndex        =   7
      Top             =   1980
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblDefensa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   1575
      TabIndex        =   6
      Top             =   1845
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Corona imperial"
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
      Height          =   255
      Left            =   1570
      TabIndex        =   5
      Top             =   1550
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image imgCross 
      Height          =   345
      Left            =   6480
      Tag             =   "1"
      Top             =   180
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Creditos:"
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
      Height          =   195
      Index           =   0
      Left            =   4690
      TabIndex        =   4
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   5505
      TabIndex        =   3
      Top             =   1545
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   5355
      TabIndex        =   2
      Top             =   1905
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmShop"
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
Private m_Number As Integer
Private m_Increment As Integer
Private m_Interval As Integer

Private Sub imgCross_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Audio.PlayWave(SND_CLICK)
imgCross.Picture = General_Load_Picture_From_Resource_Ex("_48")
imgCross.Tag = "1"

End Sub

Private Sub imgCross_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgCross.Tag = "0" Then
 imgCross.Picture = General_Load_Picture_From_Resource_Ex("_49")
    
    imgCross.Tag = "1"
End If
End Sub
 
Private Sub imgcross_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
 Unload Me
 End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Unload Me
End If

End Sub
Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource_Ex("_69")
Call FormParser.Parse_Form(Me)
Call PedirPremios

Label2.Caption = CurrentUser.Creditos
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If imgComprar.Tag = "1" Then
    imgComprar.Picture = Nothing
    imgComprar.Tag = "0"
End If

If imgVender.Tag = "1" Then
    imgVender.Picture = Nothing
    imgVender.Tag = "0"
End If

If imgCross.Tag = "1" Then
    imgCross.Picture = Nothing
    imgCross.Tag = "0"
End If

End Sub
Private Sub imgComprar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
    imgComprar.Picture = General_Load_Picture_From_Resource_Ex("_44")
    imgComprar.Tag = "1"
End Sub
Private Sub imgComprar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgComprar.Tag = "0" Then
        imgComprar.Picture = General_Load_Picture_From_Resource_Ex("_45")
        imgComprar.Tag = "1"
    End If
End Sub
Private Sub imgComprar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call Form_MouseMove(Button, Shift, X, Y)
   ' $ Shermie80 / Fix bug anti-boluditos.
   If List1.ListIndex = -1 Then
      frmMensaje.msg.Caption = "Debes seleccionar un item para comprar."
      frmMensaje.Show , frmShop
    Exit Sub
   End If
   'Fin
  

    Call WriteRPremios(List1.ListIndex + 1)
    
    If CurrentUser.Creditos >= PremiosInv(List1.ListIndex + 1).Puntos Then
        CurrentUser.Creditos = CurrentUser.Creditos - PremiosInv(List1.ListIndex + 1).Puntos
        If CurrentUser.Creditos >= PremiosInv(List1.ListIndex + 1).Puntos Then
            Label3.ForeColor = vbWhite
        Else
            Label3.ForeColor = vbWhite
        End If
        Label3.Caption = PremiosInv(List1.ListIndex + 1).Puntos
        Label2.Caption = CurrentUser.Creditos
    End If
End Sub
Private Sub imgVender_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Audio.PlayWave(SND_CLICK)
        imgVender.Picture = General_Load_Picture_From_Resource_Ex("_46")
        imgVender.Tag = "1"
End Sub
Private Sub imgVender_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgVender.Tag = "0" Then
        imgVender.Picture = General_Load_Picture_From_Resource_Ex("_47")
        imgVender.Tag = "1"
    End If
End Sub
Private Sub imgVender_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
      frmMensaje.msg.Caption = "No puedes vender este objeto."
      frmMensaje.Show , frmShop
End Sub
Private Sub list1_Click()

lblName.Caption = List1.Text
Label3.Caption = PremiosInv(List1.ListIndex + 1).Puntos
Label2.Caption = CurrentUser.Creditos

lblName.Visible = True
Label3.Visible = True
lblDefensa(0).Visible = True
lblDefensa(1).Visible = True

  Select Case List1.ListIndex + 1
    
    Case 1
      Call DrawGrhtoHdc(PicItem.hDC, 21675, 0, 0)
      lblDefensa(0).Caption = "Defensa minima: 5"
      lblDefensa(1).Caption = "Defensa maxima: 10"
       
    Case 2
      Call DrawGrhtoHdc(PicItem.hDC, 884, 0, 0)
      lblDefensa(0).Caption = "Defensa minima: 5"
      lblDefensa(1).Caption = "Defensa maxima: 10"
      
    Case 3
      Call DrawGrhtoHdc(PicItem.hDC, 32007, 0, 0)
      lblDefensa(0).Caption = "Defensa minima: 10"
      lblDefensa(1).Caption = "Defensa maxima: 10"
      
    Case 4
      Call DrawGrhtoHdc(PicItem.hDC, 311, 0, 0)
      lblDefensa(0).Caption = "Defensa minima: 10"
      lblDefensa(1).Caption = "Defensa maxima: 10"
      
    Case 5
      Call DrawGrhtoHdc(PicItem.hDC, 26195, 0, 0)
      lblDefensa(0).Caption = "Defensa minima: 15"
      lblDefensa(1).Caption = "Defensa maxima: 25"
      
    Case 6
      Call DrawGrhtoHdc(PicItem.hDC, 1717, 0, 0)
      lblDefensa(0).Caption = "Defensa minima: 15"
      lblDefensa(1).Caption = "Defensa maxima: 25"
        
  End Select
   
End Sub
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub listinv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
