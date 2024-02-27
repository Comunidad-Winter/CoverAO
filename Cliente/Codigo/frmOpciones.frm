VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6870
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
   ForeColor       =   &H00000000&
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmIdioma 
      Caption         =   "$465"
      Height          =   705
      Left            =   120
      TabIndex        =   37
      Top             =   150
      Width           =   3255
      Begin VB.ComboBox Español 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOpciones.frx":0152
         Left            =   180
         List            =   "frmOpciones.frx":015C
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Español"
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton imgConfigTeclas 
      Caption         =   "$69"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   6870
      Width           =   3255
   End
   Begin VB.Frame Frame4 
      Caption         =   "$68"
      Height          =   4065
      Left            =   3510
      TabIndex        =   9
      Top             =   3570
      Width           =   3285
      Begin VB.CheckBox chkOp 
         Caption         =   "$583"
         Height          =   285
         Index           =   10
         Left            =   180
         TabIndex        =   40
         Top             =   2040
         Width           =   2715
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$584"
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   39
         Top             =   1800
         Width           =   2715
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$84"
         Height          =   285
         Index           =   11
         Left            =   180
         TabIndex        =   36
         Top             =   1350
         Width           =   2715
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$85"
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   35
         Top             =   1080
         Width           =   2715
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$80"
         Height          =   285
         Index           =   6
         Left            =   180
         TabIndex        =   34
         Top             =   810
         Width           =   2715
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$87"
         Height          =   285
         Index           =   8
         Left            =   180
         TabIndex        =   31
         Top             =   540
         Width           =   2715
      End
      Begin VB.ListBox lstIgnore 
         Height          =   1035
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   2730
         Width           =   2895
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$88"
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   16
         Top             =   270
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "$83"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   2400
         Width           =   2925
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "$66"
      Height          =   3345
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   3285
      Begin VB.ListBox imgSkin 
         Height          =   1620
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   1350
         Width           =   2895
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$82"
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   27
         Top             =   840
         Width           =   2715
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$81"
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   15
         Top             =   570
         Width           =   2715
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$80"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label lblSkinData 
         BackStyle       =   0  'Transparent
         Caption         =   "$582"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   3030
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "$75"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   29
         Top             =   1140
         Width           =   2925
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "$69"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   5190
      Width           =   3255
      Begin VB.CommandButton cmdWeb 
         Caption         =   "$70"
         Height          =   345
         Index           =   1
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   690
         Width           =   2895
      End
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&www.coverao.com.ar"
         Height          =   345
         Index           =   0
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   300
         Width           =   2895
      End
      Begin VB.CommandButton imgSoporte 
         Caption         =   "$71"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Audio"
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   990
      Width           =   3255
      Begin VB.CheckBox chkInvertir 
         Caption         =   "$76"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   1530
         Value           =   2  'Grayed
         Width           =   2985
      End
      Begin VB.HScrollBar scrMidi 
         Height          =   315
         LargeChange     =   15
         Left            =   150
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   22
         Top             =   3570
         Width           =   2895
      End
      Begin VB.HScrollBar scrAmbient 
         Enabled         =   0   'False
         Height          =   315
         LargeChange     =   15
         Left            =   150
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   21
         Top             =   3000
         Width           =   2895
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   315
         LargeChange     =   15
         Left            =   150
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   19
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$78"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   600
         Width           =   2985
      End
      Begin VB.CheckBox chkopo 
         Caption         =   "$77"
         Height          =   285
         Left            =   180
         TabIndex        =   13
         Top             =   900
         Width           =   2955
      End
      Begin VB.CheckBox chkOp 
         Caption         =   "$79"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   2985
      End
      Begin VB.TextBox txtMidi 
         Height          =   285
         Left            =   2385
         TabIndex        =   1
         Top             =   1845
         Width           =   345
      End
      Begin VB.CheckBox chkMidi 
         Caption         =   "$91"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   1230
         Width           =   2985
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "$72"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   24
         Top             =   3360
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "$73"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   2790
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "$74"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   2190
         Width           =   2835
      End
      Begin VB.Label lblNextMidi 
         Caption         =   "»"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2760
         TabIndex        =   11
         Top             =   1875
         Width           =   135
      End
      Begin VB.Label lblBackMidi 
         Caption         =   "«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2265
         TabIndex        =   10
         Top             =   1875
         Width           =   135
      End
      Begin VB.Label lblMidi 
         BackStyle       =   0  'Transparent
         Caption         =   "$92"
         Height          =   255
         Left            =   195
         TabIndex        =   8
         Top             =   1875
         Width           =   2055
      End
   End
   Begin VB.CommandButton imgSalir 
      Caption         =   "$25"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   7260
      Width           =   3255
   End
   Begin VB.Menu mnuIgnore 
      Caption         =   "Ignorar"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuQuitarIgnorado 
         Caption         =   "Quitar"
      End
      Begin VB.Menu mnuAgregarIgnorado 
         Caption         =   "Agregar"
      End
   End
End
Attribute VB_Name = "frmOpciones"
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
 
Private bLoading As Boolean


Private Sub chkMidi_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If chkMidi.Value = 1 Then 'Si lo desactivo
    txtMidi.Enabled = False
    lblNextMidi.Enabled = False
    lblBackMidi.Enabled = False
    
    If CurrentUser.Logged Then 'Y tengo la música activada
        If Musica Then
            Audio.MusicActivated = True
            scrMidi.Enabled = True
            scrMidi.Value = Audio.MusicVolume
            '0(MapDat.music_number)
            'Audio.MusicActivated = True
            'scrMidi.Enabled = True
            'scrMidi.Value = Audio.MusicVolume
        End If
    End If

Else
    txtMidi.Enabled = True
    lblNextMidi.Enabled = True
    lblBackMidi.Enabled = True
End If


End Sub


Private Sub chkop_Click(Index As Integer)

Dim Opcion As Byte


Opcion = IIf(chkOp(Index).Value = vbChecked, 1, 0)

Select Case Index

Case 9 'Habilita mensajes globales

HabilitarMensajesGlobales = Not HabilitarMensajesGlobales


Case 2 'Nombre de mapa

Dim map_x As Byte
Dim map_y As Byte

 

If VerLugar = 0 Then
    VerLugar = 1
    frmMain.Label2(0).Caption = Map_Name_Get
Else
    VerLugar = 0
    Call Char_MapPosGet(CurrentUser.UserCharIndex, map_x, map_y)
    frmMain.Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.UserMap & ", " & map_x & ", " & map_y
End If
 

Case 8

NickModerno = CBool(Opcion)

Call CargarMedidasNombresModernos

Case 32
Call Audio.PlayWave(SND_CLICK)


        If MsgBox(Locale_GUI_Frase(338), vbQuestion + vbYesNo, Locale_GUI_Frase(339)) = vbYes Then
            If CursorHabilitado = 1 Then
                CursorHabilitado = 0
            Else
                CursorHabilitado = 1
            End If
        End If
        
Call WriteVar(App.Path & "\Init\CovAoInit.ini", "CONFIG", "CursorHabilitado", CursorHabilitado)


Case 0


    
    Musica = Not Musica
            
    If Not Musica Then
        Audio.MusicActivated = False
        scrMidi.Enabled = False
    Else
        If Not Audio.MusicActivated Then  'Prevent the music from reloading
            Audio.MusicActivated = True
            scrMidi.Enabled = True
         '   scrMidi.Value = Audio.MusicVolume
        End If
    End If
    
 
    
    Case 1
    
       

    SonidoHabilitado = Not SonidoHabilitado
    
    If Not SonidoHabilitado Then
        Audio.SoundActivated = False
        RainBufferIndex = 0
        frmMain.IsPlaying = PlayLoop.plNone
        scrVolume.Enabled = False
    Else
        Audio.SoundActivated = True
        scrVolume.Enabled = True
        scrVolume.Value = Audio.SoundVolume
        
    End If
    
    
    
    
    Case 4
    
    
    'If CantidadEnMacros Then
    'UpdateMacroLabels (1)
    'Else
    'UpdateMacroLabels (0)
    'End If
    
    chkOp(Index).Value = IIf((CantidadEnMacros), 1, 0)
 
    Case 10
    
    BloqueoAlCaminar = Me.chkOp(Index).Value
    
    Me.chkOp(Index).Value = IIf(BloqueoAlCaminar, 1, 0)
 
 
    
End Select
End Sub

Private Sub cmdWeb_Click(Index As Integer)
Select Case Index
Case 0
ShellExecute Me.hwnd, "open", Chr$(34) & "http://www.coverao.com.ar/" & Chr$(34), vbNullString, vbNullString, 1

Case 1
MsgBox "No hay función para este botón"
End Select

End Sub

Private Sub imgSkin_click()

                 NombreSkin = imgSkin.list(imgSkin.ListIndex)
                
                Call WriteVar(App.Path & "\Init\CovAoInit.ini", "LAUNCHER", "NombreSkin", NombreSkin)
                MsgBox "Para que los cambios en esta opción sean reflejados, deberá reiniciar el cliente.", vbQuestion, "Advertencia"
                
End Sub

Private Sub imgConfigTeclas_Click()
frmCustomKeys.Show vbModeless, frmOpciones
End Sub
Private Sub imgSalir_Click()
     Call Audio.PlayWave(SND_CLICK)
    
 
    
    Me.Visible = False
  
    
End Sub

Private Sub imgSoporte_Click()
    Me.Visible = False
    frmGMAyuda.Show vbModeless, frmMain
    frmGMAyuda.TxtSoporte.SetFocus
End Sub



Private Sub Form_Load()

Call FormParser.Parse_Form(Me)

End Sub

 
Public Sub AgregarIgnorado(ByVal Nick As String)

On Error Resume Next

Dim i As Long

Nick = UCase$(Nick)

For i = 0 To lstIgnore.ListCount
    If UCase$(lstIgnore.list(i)) = Nick Then
        
        Call lstIgnore.RemoveItem(i)
        
        If CurrentUser.Logged Then
            Call AddtoRichTextBox(Nick & " " & Locale_GUI_Frase(262), 0, 0, 0, 0, 0, 0, 8)
        Else
            Call MensajeAdvertencia(Nick & " " & Locale_GUI_Frase(262))
        End If
        
        Exit Sub
    End If
Next i

lstIgnore.AddItem Nick
If CurrentUser.Logged Then Call AddtoRichTextBox(Nick & " " & Locale_GUI_Frase(263), 0, 0, 0, 0, 0, 0, 8)

End Sub

 
Public Sub Init()

On Error Resume Next

Dim t() As String, i As Integer, file_name As String, tBtArr() As Byte

bLoading = True

Español.ListIndex = 0

chkOp(2).Value = VerLugar 'ver mapa
chkOp(3).Value = Nombres 'ver nombre de jugadores
chkOp(4).Value = CantidadEnMacros 'Macro con numeros
chkOp(5).Value = 0 'cursores graficos
chkOp(6).Value = 0 'publicidad adulta
chkOp(7).Value = 0 'Ver dialogos en consola
chkOp(8).Value = IIf((NickModerno = True), 1, 0) 'Nombres modernos
chkOp(9).Value = IIf((HabilitarMensajesGlobales = True), 1, 0) 'Mensajes globales
chkOp(10).Value = BloqueoAlCaminar 'No moverse al hablar
chkOp(11).Value = 0 'Chat faccionario


If lstIgnore.ListCount = 0 Then
    t = Split(ListaIgnorados, "¬")
    
    lstIgnore.Clear
    
    For i = 0 To UBound(t)
        lstIgnore.AddItem t(i)
    Next i
End If


imgSkin.Clear

file_name = Dir$(App.Path & "\Skins\")
Do While Len(file_name) > 0
    If Not _
        (file_name = ".") Or _
        (file_name = "..") Or _
        (Right$(file_name, 3) <> "ias") _
    Then
            If Resource_File_Exists(App.Path & "\Skins\" & file_name, "todo.jpg") Then
                imgSkin.AddItem mid$(file_name, 1, Len(file_name) - 4)
            End If
    End If
    
    file_name = Dir$()
Loop



 

bLoading = False
    Me.Show vbModeless, frmMain
End Sub

Private Sub lstIgnore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    mnuQuitarIgnorado.Enabled = (lstIgnore.ListIndex <> -1)
    PopupMenu mnuIgnore
End If

End Sub

Private Sub mnuAgregarIgnorado_Click()

Dim Resp As String
Resp = InputBox("Escriba el nombre del usuario que desea ignorar (también puede usar el comando /IGNORAR nick)", "Ignorar usuario")
If Resp <> vbNullString Then Call AgregarIgnorado(Resp)

End Sub

Private Sub mnuQuitarIgnorado_Click()

If lstIgnore.ListIndex = -1 Then Exit Sub
lstIgnore.RemoveItem lstIgnore.ListIndex

End Sub


Private Sub scrMidi_Change()
If Musica <> False Then
    Audio.MusicVolume = scrMidi.Value
    MusicVolume = Audio.MusicVolume
End If
End Sub

Private Sub scrVolume_Change()

If SonidoHabilitado = 1 Then
    Audio.SoundVolume = scrVolume.Value
    FXVolume = Audio.SoundVolume
End If

End Sub

Private Sub txtMidi_Change()

If Musica = False Then Exit Sub

If val(txtMidi.Text) > 0 And (val(txtMidi.Text) <> Audio.MusicActual) Then
    'If Not Sound.Music_Load(val(txtMidi.Text), Sound.VolumenActualMusic) Then
    '    txtMidi.Text = Sound.MusicActual
    'Else
        'Sound.Music_Stop
        Call Audio.PlayMIDI(txtMidi.Text)
    'End If
End If

End Sub
