VERSION 5.00
Begin VB.Form frmPregunta 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPregunta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label msg 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3675
      WordWrap        =   -1  'True
   End
   Begin VB.Image img 
      Height          =   465
      Index           =   1
      Left            =   2085
      Tag             =   "1"
      Top             =   2070
      Width           =   1215
   End
   Begin VB.Image img 
      Height          =   465
      Index           =   0
      Left            =   615
      Tag             =   "1"
      Top             =   2070
      Width           =   1200
   End
End
Attribute VB_Name = "frmPregunta"
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

Private Enum eAction
    DestruirObjetos = 1
    Retirar = 2
    frmMuerto = 3
    LinkWeb = 4
    EstablecerHogar = 5
    BorrarPersonaje = 6
    Desactualizado = 7
End Enum

Public Accion As Byte

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1

Private NickAux As String

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long


Public Sub SetAccion(ByVal action As Byte, Optional ByRef a As String = vbNullString)

    On Error GoTo SetAccion_Err
    
    Accion = action
    
    Select Case action
    
    Case 1 'Destroy OBJ
        Me.msg = Locale_GUI_Frase(542) '¿Está seguro que desea tirar el objeto seleccionado? Esta acción DESTRUIRÁ EL OBJETO.
    
    Case 2 '/Retirar
        Me.msg = Locale_Parse_Pregunta(1) 'Los renegados no pueden formar parte de grupos y pierden 10% de la experiencia ganada, para recuperar su ciudadanía deberá pagar un precio en experiencia luego. En caso de entender esto, presione "Aceptar".
    
    Case 3 'Cartel muerto
        Me.msg = Locale_GUI_Frase(548) 'Has muerto. Presiona aceptar para liberar tu cuerpo y ser revivido.
    
    Case 4 'Borrar PJ
        Me.msg = Locale_GUI_Frase(645) & " " & a & " " & Locale_GUI_Frase(646)
    
    Case 5 'Nuevo hogar
        Me.msg = Locale_GUI_Frase(598) & " " & Map_Name_Get & " " & Locale_GUI_Frase(599) '¿Deseas que Ciudad sea tu nuevo hogar?
    
    Case 6
    
    Case 7 'Juego desactualizado
        Me.msg = Locale_Error(48)
    
    Case 8 'Casamiento
        Me.msg = Locale_Parse_Pregunta(6, a) '
        NickAux = a

    End Select
    
    Exit Sub

SetAccion_Err:
     Call RegistrarError(Err.number, Err.Description, "frmPregunta.SetAccion", Erl)
     Resume Next
End Sub

Private Sub img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Audio.PlayWave(SND_CLICK)


Select Case Index

Case 0
img(0).Picture = General_Load_Picture_From_Resource_Ex("_12")
img(0).Tag = "1"
Case 1

img(1).Picture = General_Load_Picture_From_Resource_Ex("_10")
img(1).Tag = "1"

End Select
 

End Sub

Private Sub img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index

Case 0
If img(0).Tag = "0" Then
img(0).Picture = General_Load_Picture_From_Resource_Ex("_11")
img(0).Tag = "1"
End If

Case 1
If img(1).Tag = "0" Then
    img(1).Picture = General_Load_Picture_From_Resource_Ex("_66")
    img(1).Tag = "1"
End If

End Select


End Sub

Private Sub img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


        Call Form_MouseMove(Button, Shift, X, Y)
        
        Select Case Index
        
        Case 0 'Aceptar
        
            Select Case Accion
            
            Case 1
                If Not ClientTCP.DeadCheck Then
                        Call WriteDropDestroy(Inventario.SelectedItem, CantidadGlobal)
                End If
                
            Case 2
                If Not ClientTCP.DeadCheck Then
                    Call WriteRetirarFaccion
                End If
            Case 3
                If ClientTCP.DeadCheck Then
                    Call WriteRegresarHogar
                End If
                
            Case 4 'Borrar PJ

                If IntervaloPermiteConectar Then
                    Call FormParser.Parse_Form(frmPregunta, E_WAIT)
                    
                    If Len(frmCharList.lblAccData(frmCharList.intSelChar).Caption) <= 0 Then
                        MsgBox "Selecciona un personaje para eliminar"
                        Exit Sub
                    End If
                    Cuenta.UserCode = CStr(frmCharList.lblAccData(frmCharList.intSelChar).Caption)
                    Cuenta.EsChange = 2
                    
                    If frmMain.Socket1.Connected Then
                        EstadoLogin = E_MODO.BorrarPersonaje
                        Call Login
                        
                    Else
                        EstadoLogin = E_MODO.BorrarPersonaje
                        frmMain.Socket1.HostName = CurServerIP
                        frmMain.Socket1.RemotePort = CurServerPort
                        frmMain.Socket1.Connect
                        
                    End If
                    
                End If
    

                        
            Case 5
                If Not ClientTCP.DeadCheck Then
                    Call WriteSeleccionarHogar(1)
                End If

            Case 7 'Desactualizado
            Call CloseClient(True, True)
            
            Case 8 'Acepta casamiento
                Call WriteCasamiento(NickAux, 0)
                
            End Select
                
        Case 1 'Salir
        
        End Select
        
        
        Accion = 0
        Me.msg.Caption = vbNullString
        NickAux = vbNullString
        Me.Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Me.msg.Caption = vbNullString
    Me.Visible = False
End If

End Sub
Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource_Ex("_9")
Call FormParser.Parse_Form(Me)
Make_Transparent_Form Me.hwnd, 210
Call Audio.PlayWave(SND_INFO)

    
    'Con esto el formulario queda abierto en cualquier pestaña hasta que el usuario la cierre.
    'Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
      
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If img(0).Tag = "1" Then
    img(0).Picture = Nothing
    img(0).Tag = "0"
End If

If img(1).Tag = "1" Then
    img(1).Picture = Nothing
    img(1).Tag = "0"
End If

End Sub

Private Sub msg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Me.Visible = False
End If

End Sub
Private Sub msg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
 
