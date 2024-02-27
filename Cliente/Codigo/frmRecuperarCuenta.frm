VERSION 5.00
Begin VB.Form frmRecuperarCuenta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$585"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4635
   Icon            =   "frmRecuperarCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "$585"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox txtUserCode 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   720
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   720
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CheckBox ChPass 
      Caption         =   "$589"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$586"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$588"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   495
      TabIndex        =   6
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$587"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "frmRecuperarCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cerrar()
    If frmMain.Socket1.State <> sckClosed Then frmMain.Socket1.Disconnect
    Unload Me
End Sub

Private Sub ChPass_Click()

    If ChPass.Value = 1 Then
        txtPassword.PasswordChar = vbNullString
        txtUserCode.PasswordChar = vbNullString
    Else
        txtPassword.PasswordChar = "*"
        txtUserCode.PasswordChar = "*"
    End If

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Cerrar
End Sub

Private Sub cmdAceptar_click()

    If IntervaloPermiteConectar Then
                    
        If CheckDat = False Then Exit Sub

        Cuenta.UserAccount = txtNombre.Text
        Cuenta.UserCode = txtUserCode.Text
        Cuenta.UserPassword = txtPassword.Text
        Cuenta.EsChange = 0
        cmdAceptar.Caption = Locale_GUI_Frase(644) 'Conectando al servidor...
        cmdAceptar.Enabled = False
        
        Call FormParser.Parse_Form(Me, E_WAIT)
        
        If frmMain.Socket1.Connected Then
            EstadoLogin = E_MODO.RecuperarCuenta
            Call Login
            Exit Sub

        Else
            EstadoLogin = E_MODO.RecuperarCuenta
            frmMain.Socket1.HostName = CurServerIP
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
            
        End If
        
    End If
    
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(frmConnect)
Call FormParser.Parse_Form(Me)

Me.txtNombre.Text = vbNullString
Me.txtPassword.Text = vbNullString
Me.txtUserCode.Text = vbNullString

End Sub

Private Sub txtUserCode_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Function CheckDat() As Boolean

If txtNombre.Text = vbNullString Then
    MsgBox Locale_Error(52) 'Nombre invalido
    CheckDat = False
    Exit Function
End If

If txtPassword.Text = vbNullString Then
    MsgBox Locale_Error(49) ' Debes ingresar una contraseña válida
    CheckDat = False
    Exit Function
End If

If txtUserCode.Text = vbNullString Then
    MsgBox Locale_Error(51) 'Debes ingresar un pin válido
    CheckDat = False
    Exit Function
End If

CheckDat = True

Exit Function

End Function

