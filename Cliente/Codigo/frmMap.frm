VERSION 5.00
Begin VB.Form frmMap 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Shape1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      FillColor       =   &H001274FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   5
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Shape1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      FillColor       =   &H001274FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   4
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Shape1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      FillColor       =   &H001274FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   3
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Shape1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      FillColor       =   &H001274FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   2
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Shape1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      FillColor       =   &H001274FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   0
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Shape1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      FillColor       =   &H001274FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   1
      Left            =   3720
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   5160
      Top             =   3420
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Index           =   5
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Index           =   4
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape shpMap 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Image imgMap 
      Appearance      =   0  'Flat
      Height          =   8655
      Left            =   180
      Top             =   180
      Width           =   11655
   End
   Begin VB.Label lblMMAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Posición cursor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   1440
      TabIndex        =   1
      Top             =   8520
      Width           =   1545
   End
   Begin VB.Label lbIMAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Posición:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   1440
      TabIndex        =   0
      Top             =   8280
      Width           =   855
   End
   Begin VB.Line ln2 
      BorderColor     =   &H00694843&
      X1              =   787
      X2              =   787
      Y1              =   589
      Y2              =   12
   End
   Begin VB.Line ln1 
      BorderColor     =   &H00694843&
      X1              =   786
      X2              =   786
      Y1              =   589
      Y2              =   12
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XX As Integer, YY As Integer
Dim Valor As Byte
Private Sub Form_Load()

    Me.Picture = General_Load_Picture_From_Resource_Ex("15567")
    Call FormParser.Parse_Form(Me, e_normal)
     Me.Top = frmMain.Top
    Me.Left = frmMain.Left
 
Make_Transparent_Form Me.hwnd, 237
End Sub
Private Sub imgMap_Click()
 Image1.Top = Label2.Top
Image1.Left = Label2.Left
Image1.Visible = True
 End Sub
Private Sub imgMap_DblClick()
        Unload Me
    frmMain.SetFocus
     SetMapPoint
     
End Sub


Private Sub imgMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    X = X / 15 - 12
    Y = Y / 15 - 12
    XX = X / 27.5 + 1
    YY = Y / 25 + 1
                
     
    If XX <= 0 Or YY <= 0 Or XX > 30 Or YY > 23 Then Exit Sub
    lblMMAP.Caption = "Posición cursor: " & MapNames(MapTable(XX, YY)) & " (" & MapTable(XX, YY) & ")"
    
 
 Label2.Left = X + 13
  Label2.Top = Y + 10
End Sub

Private Sub imgMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Button = vbRightButton Then
        Unload Me
       frmMain.SetFocus
     
    End If
     
End Sub
Public Sub SetMapPoint()
    Dim i As Long, ii As Long
     For i = 1 To 30
        For ii = 1 To 23
            If MapTable(i, ii) = CurrentUser.UserMap Then
                shpMap.Top = (ii - 1) * 25 + 12 + 18
                shpMap.Left = (i - 1) * 26 + 12 + 27
                
                lbIMAP.Caption = "Posición: " & MapNames(MapTable(i, ii)) & " (" & MapTable(i, ii) & ")"
                Exit For
            End If
        Next ii
    Next i
End Sub
Public Sub SetMapPoint2(Optional ByVal numUser As Byte, Optional ByVal Mapa As Integer)
 Dim J As Long, jj As Long
    For J = 1 To 30
        For jj = 1 To 23
            If MapTable(J, jj) = Mapa Then
                Shape1(numUser).Top = (jj - 1) * 25 + 12 + 18
                Shape1(numUser).Left = (J - 1) * 26 + 12 + 27
                Exit For
            End If
        Next jj
    Next J
 
 
End Sub
Private Sub shape1_click(Index As Integer)
 If Valor = 0 Then
Label1(Index).Visible = True
 
 Valor = 1
Else
 Label1(Index).Visible = False
 Valor = 0
 End If
End Sub
 Private Sub image1_dblclick()
 Call WriteWarpChar("YO", MapTable(XX, YY), 50, 50)
         Unload Me
    frmMain.SetFocus
     SetMapPoint
     
 End Sub
