VERSION 5.00
Begin VB.Form frmGMPanel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Panel GM"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPanelGMs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "RMSG"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4320
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Usuario manual (Solo 'Mensaje usuario')"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MSG All"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   4320
      Width           =   735
   End
   Begin VB.ListBox txtmensaje 
      Height          =   2010
      Left            =   120
      TabIndex        =   5
      Top             =   870
      Width           =   4560
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "MSG User"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      Height          =   1035
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3120
      Width           =   4575
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "&Actualiza"
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboListaUsus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3675
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   110
      X2              =   4680
      Y1              =   4290
      Y2              =   4290
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   4680
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Menu mnuUsuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
      Begin VB.Menu mnuIra 
         Caption         =   "Ir al usuario"
      End
      Begin VB.Menu mnuTraer 
         Caption         =   "Traer el usuario"
      End
      Begin VB.Menu mnuInvalida 
         Caption         =   "Inválida"
      End
      Begin VB.Menu mnuManual 
         Caption         =   "Manual/FAQ"
      End
   End
   Begin VB.Menu mnuChar 
      Caption         =   "Personaje"
      Begin VB.Menu cmdAccion 
         Caption         =   "Echar"
         Index           =   0
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Sumonear"
         Index           =   2
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ir a"
         Index           =   3
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ubicación"
         Index           =   6
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Desbanear"
         Index           =   12
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "IP del personaje"
         Index           =   13
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Revivir"
         Index           =   21
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Advertir"
         Index           =   22
      End
      Begin VB.Menu cmdPremium 
         Caption         =   "Premium"
         Begin VB.Menu mnuPrem 
            Caption         =   "Info"
            Index           =   42
         End
         Begin VB.Menu mnuPrem 
            Caption         =   "Donador"
            Index           =   43
         End
         Begin VB.Menu mnuPrem 
            Caption         =   "Otorgar creditos"
            Index           =   44
         End
      End
      Begin VB.Menu cmdBanMenu 
         Caption         =   "Banear"
         Begin VB.Menu mnuBan 
            Caption         =   "Personaje"
            Index           =   1
         End
         Begin VB.Menu mnuBan 
            Caption         =   "Personaje e IP"
            Index           =   19
         End
         Begin VB.Menu mnuBan 
            Caption         =   "Clan"
            Index           =   51
         End
      End
      Begin VB.Menu mnuEncarcelar 
         Caption         =   "Encarcelar"
         Begin VB.Menu mnuCarcel 
            Caption         =   "5 Minutos"
            Index           =   5
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "15 Minutos"
            Index           =   15
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "30 Minutos"
            Index           =   30
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "Definir otro"
            Index           =   60
         End
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Información"
         Begin VB.Menu mnuAccion 
            Caption         =   "General"
            Index           =   8
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Inventario"
            Index           =   9
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Skills"
            Index           =   10
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Bóveda"
            Index           =   18
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Stats"
            Index           =   70
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Oro en banco"
            Index           =   71
         End
      End
      Begin VB.Menu mnuSilenciar 
         Caption         =   "Silenciar"
         Begin VB.Menu mnuSilencio 
            Caption         =   "5 Minutos"
            Index           =   5
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "15 Minutos"
            Index           =   15
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "30 Minutos"
            Index           =   30
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "Definir otro"
            Index           =   60
         End
      End
      Begin VB.Menu MnuChange 
         Caption         =   "Cambiar contraseña"
         Index           =   62
      End
      Begin VB.Menu mnuIrCerca 
         Caption         =   "Ir cerca"
         Index           =   64
      End
   End
   Begin VB.Menu cmdHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Insertar comentario"
         Index           =   4
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Enviar hora"
         Index           =   5
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Enemigos en mapa"
         Index           =   7
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Limpiar Mapa"
         Index           =   15
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Activar/Desactivar Centinela"
         Index           =   16
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios trabajando"
         Index           =   23
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios ocultandose"
         Index           =   24
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Bloquear tile"
         Index           =   26
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios en el mapa"
         Index           =   30
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "GMs Conectados"
         Index           =   72
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Clan Conectados"
         Index           =   73
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Remover estupidez"
         Index           =   74
      End
      Begin VB.Menu IP 
         Caption         =   "Direcciónes de IP"
         Index           =   0
         Begin VB.Menu mnuIP 
            Caption         =   "Buscar IP's Coincidentes"
            Index           =   14
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Banear una IP"
            Index           =   17
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Lista de IPs baneadas"
            Index           =   25
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Reloads BanIPs"
            Index           =   45
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Unban IP"
            Index           =   52
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Last IP"
            Index           =   53
         End
      End
   End
   Begin VB.Menu Admin 
      Caption         =   "Administración"
      Index           =   0
      Begin VB.Menu mnuAdmin 
         Caption         =   "Apagar servidor"
         Index           =   27
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Grabar personajes"
         Index           =   28
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Visualizar consultas"
         Index           =   29
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Buscar NPCs/Objetos"
         Index           =   33
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Limpiar el mundo"
         Index           =   34
      End
      Begin VB.Menu mnuRecargar 
         Caption         =   "Actualizar"
         Index           =   35
         Begin VB.Menu mnuReload 
            Caption         =   "Objetos"
            Index           =   1
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Hechizos"
            Index           =   2
         End
         Begin VB.Menu mnuReload 
            Caption         =   "NPCs"
            Index           =   3
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Sockets"
            Index           =   4
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Puntual server.ini"
            Index           =   5
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Echar todos los PJs"
            Index           =   6
         End
      End
      Begin VB.Menu Ambiente 
         Caption         =   "Habilitar/deshabilitar servidor solo GMs"
         Index           =   31
      End
      Begin VB.Menu Ambiente 
         Caption         =   "Estado climático"
         Index           =   32
      End
      Begin VB.Menu mnuPwChars 
         Caption         =   "Cambiar contraseña PJs"
      End
      Begin VB.Menu cmdOnlineClan 
         Caption         =   "Miembros Clan"
         Index           =   57
      End
      Begin VB.Menu cmdOnlineClan 
         Caption         =   "Rajar Miembro Clan"
         Index           =   61
      End
      Begin VB.Menu MenuObj 
         Caption         =   "Objetos"
         Begin VB.Menu mnuObject 
            Caption         =   "Crear OBJ #Num"
            Index           =   58
         End
         Begin VB.Menu mnuObject 
            Caption         =   "Destruir objeto"
            Index           =   59
         End
         Begin VB.Menu mnuObject 
            Caption         =   "Destruir objetos area"
            Index           =   60
         End
      End
   End
   Begin VB.Menu mnuOtros 
      Caption         =   "Otros"
      Begin VB.Menu mnuCr 
         Caption         =   "Cuenta regresiva"
         Begin VB.Menu mnuCre 
            Caption         =   "5"
            Index           =   36
         End
         Begin VB.Menu mnuCre 
            Caption         =   "10"
            Index           =   37
         End
         Begin VB.Menu mnuCre 
            Caption         =   "Otro"
            Index           =   38
         End
      End
      Begin VB.Menu mnuParticula 
         Caption         =   "Particula"
         Index           =   39
      End
      Begin VB.Menu mnuMultiplicador 
         Caption         =   "Multiplicadores"
         Begin VB.Menu mnuMulti 
            Caption         =   "Oro"
            Index           =   40
         End
         Begin VB.Menu mnuMulti 
            Caption         =   "Experiencia"
            Index           =   41
         End
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "NPCs"
         Begin VB.Menu mnuNPC 
            Caption         =   "NPC dialogo"
            Index           =   20
         End
         Begin VB.Menu mnuNPC 
            Caption         =   "Atraer NPC"
            Index           =   46
         End
         Begin VB.Menu mnuNPC 
            Caption         =   "Reset inventario "
            Index           =   47
         End
         Begin VB.Menu mnuNPC 
            Caption         =   "Crear NPC #Num"
            Index           =   48
         End
         Begin VB.Menu mnuNPC 
            Caption         =   "Kill NPC"
            Index           =   49
         End
         Begin VB.Menu mnuNPC 
            Caption         =   "Kill NPCs Area"
            Index           =   50
         End
         Begin VB.Menu mnuNPC 
            Caption         =   "Kill NPC/Respawn"
            Index           =   68
         End
         Begin VB.Menu mnuNPC 
            Caption         =   "Crear NPC con Respawn"
            Index           =   77
         End
      End
      Begin VB.Menu mnuMaps 
         Caption         =   "Teleport/Mapas"
         Begin VB.Menu mnuTelep 
            Caption         =   "Crear Portal"
            Index           =   54
         End
         Begin VB.Menu mnuTelep 
            Caption         =   "Destruir Portal"
            Index           =   55
         End
         Begin VB.Menu mnuTelep 
            Caption         =   "Cambiar status Mapa"
            Index           =   56
         End
         Begin VB.Menu mnuTelep 
            Caption         =   "Teletransportar User Mapa X Y"
            Index           =   65
         End
      End
      Begin VB.Menu mnuNuevos 
         Caption         =   "Mostrar/Ocultar Nick"
         Index           =   63
      End
      Begin VB.Menu mnuNuevos 
         Caption         =   "Invisible YO"
         Index           =   67
      End
      Begin VB.Menu mnuNuevos 
         Caption         =   "Modificar personaje"
         Index           =   69
      End
      Begin VB.Menu mnuNuevos 
         Caption         =   "Respawn List"
         Index           =   75
      End
      Begin VB.Menu mnuNuevos 
         Caption         =   "Trigger en Pos"
         Index           =   76
      End
      Begin VB.Menu mnuNuevos 
         Caption         =   "/NAVE"
         Index           =   78
      End
      Begin VB.Menu mnuNuevos 
         Caption         =   "Guardar Mapa"
         Index           =   79
      End
   End
End
Attribute VB_Name = "frmGMPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Dim Nick As String
 

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible _
    Lib "user32" ( _
        ByVal hwnd As Long) As Long

'Esta función retorna el número de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength _
    Lib "user32" _
    Alias "GetWindowTextLengthA" ( _
        ByVal hwnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenará el texto devuelto, y el Lenght de la cadena en el último parámetro
'que obtuvimos con el Api GetWindowTextLength
Private Declare Function GetWindowText _
    Lib "user32" _
    Alias "GetWindowTextA" ( _
        ByVal hwnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long

'Esta es la función Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow _
    Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&
Public CANTv As Byte
Public Function Listar() As String
Static alter As String
Dim buf As Long, handle As Long, titulo As String, lenT As Long, ret As Long
    'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
    handle = GetWindow(hwnd, GW_HWNDFIRST)

    'Este bucle va a recorrer todas las ventanas.
    'cuando GetWindow devielva un 0, es por que no hay mas
    Do While handle <> 0
        'Tenemos que comprobar que la ventana es una de tipo visible
        If IsWindowVisible(handle) Then
            'Obtenemos el número de caracteres de la ventana
            lenT = GetWindowTextLength(handle)
            'si es el número anterior es mayor a 0
            If lenT > 0 Then
                'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
                titulo = String$(lenT, 0)
                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                'y también debemos pasarle el Hwnd de dicha ventana
                ret = GetWindowText(handle, titulo, lenT + 1)
                titulo$ = Left$(titulo, ret)
                'La agregamos al ListBox
                Listar = titulo & "#" & Listar
                CANTv = CANTv + 1
            End If
        End If
        'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
        handle = GetWindow(handle, GW_HWNDNEXT)
       Loop
End Function

Private Sub Ambiente_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub Check1_Click()
 If Check1.Value = 1 Then
  cboListaUsus.Visible = False
  Text1.Visible = True
 Else
  cboListaUsus.Visible = True
  Text1.Visible = False
 End If
End Sub

Private Sub cmdAccion_Click(Index As Integer)

Dim reason As String
Dim tmp As String
Dim CR, CR2 As Byte

Nick = cboListaUsus.Text

Select Case Index

Case 0 '/ECHAR nick0
If LenB(Nick) <> 0 Then Call WriteKick(Nick)

Case 1 '/ban motivo@nick

    If MsgBox("¿Está seguro que desea banear al personaje " & cboListaUsus.Text & "?", vbYesNo + vbQuestion) = vbYes Then
        Call WriteBanChar(Nick, 1)
    End If
    
Case 2 '/sum nick 0
    If LenB(Nick) <> 0 Then Call WriteSummonChar(Nick)
Case 3 '/ira nick0
    If LenB(Nick) <> 0 Then Call WriteGoToChar(Nick)
Case 4 '/rem
    tmp = InputBox("¿Comentario?", "Ingrese comentario")
    Call WriteComment(tmp)
Case 5 '/hora
    Call AddtoRichTextBox("La hora en el mundo es: " & tHora & ":" & tMinuto, 0, 0, 0, 0, 0, 0, 8)
    'Call Protocol.WriteServerTime
Case 6 '/donde nick 0
    If LenB(Nick) <> 0 Then Call WriteWhere(Nick)
Case 7 '/nene
    tmp = InputBox("¿En qué mapa?", vbNullString)
    Call WriteCreaturesInMap(tmp)
Case 8 '/info nick
    Call WriteRequestCharInfo(Nick)
Case 9 '/inv nick
    Call WriteRequestCharInventory(Nick)
Case 10 '/skills nick
    Call WriteRequestCharSkills(Nick)
Case 11 '/carcel minutos nick
   tmp = InputBox("¿Minutos a encarcelar? (hasta 60)", vbNullString)
    
    If MsgBox("¿Esta seguro que desea encarcelar al personaje " & Nick & "?", vbYesNo + vbQuestion) = vbYes Then
       ' Call Clienttcp.ParseUserCommand("/CARCEL " & nick & "@" & "Mal comportamiento." & "@" & tmp)
    End If
Case 12 '/unban nick0
    If MsgBox("¿Esta seguro que desea removerle el ban al personaje " & Nick & "?", vbYesNo + vbQuestion) = vbYes Then
      Call WriteBanChar(Nick, 0)
    End If
Case 13 '/nick2ip nick 0
    Call WriteNickToIP(Nick)
Case 14 '/IP2NICK nick
    tmp = InputBox("Ingrese IP a relacionar con PJs", vbNullString)
    Call WriteIPToNick(str2ipv4l(tmp))
Case 15
    tmp = InputBox("¿Mapa?", vbNullString)
    MsgBox "Limpiar mapa"
Case 16
    MsgBox "activar deasc centinela"
Case 17
     tmp = InputBox("Escriba el ip o el nick del PJ.", "Banear IP")
     reason = InputBox("Escriba el motivo del ban.", "Banear IP")
    If MsgBox("¿Esta seguro que desea banear la IP del personaje " & Nick & "?", vbYesNo + vbQuestion) = vbYes Then
    Call ClientTCP.ParseUserCommand("/BANIP " & tmp & " " & "BANIP")  'We use the Parser to control the command format
    End If
Case 18 '/bov nick
    Call WriteRequestCharBank(Nick)
Case 19
     tmp = InputBox("Escriba el ip o el nick del PJ.", "Banear IP")
     reason = InputBox("Escriba el motivo del ban.", "Banear IP")
    If MsgBox("¿Esta seguro que desea banear la IP del personaje " & Nick & "?", vbYesNo + vbQuestion) = vbYes Then
   ' Call Clienttcp.ParseUserCommand("/BANIP " & tmp & " " & "BANIP")  'We use the Parser to control the command format
    End If
Case 20 'slot sin usar
    MsgBox "talkas"
Case 21 '/revivir nick0
    Call WriteReviveChar(Nick)
Case 22
    MsgBox "advertir"
Case 23 'Trabajando
    Call WriteWorking
Case 24 '
    MsgBox "Usuarios ocultand"
Case 25 'Lista Ip
    Call WriteBannedIPList
Case 26 'Bloq
    Call WriteTileBlockedToggle
Case 27 'Apagar
    Call WriteTurnOffServer
Case 28 'Pjs
    Call WriteSaveChars
Case 29 'bakcup
    MsgBox "show sos"
Case 30 'OnlineMap
    Call WriteOnlineMap(CurrentUser.UserMap)
Case 31 '
MsgBox "solo gms"
Case 32
tmp = InputBox("1; LLuvia, 2; LLuvia electrica, 3; Nieve, 0; Nada", vbNullString)
Call WriteRainToggle(tmp)

Case 33
    MsgBox "buscar nc objetos"
Case 34 'Limpiar mundo
    MsgBox "limpiar mundo"
Case 35 '/silencio minutos nick
    tmp = InputBox("¿Minutos a silenciar? (hasta 60)", vbNullString)
    If MsgBox("¿Esta seguro que desea silenciar al personaje " & Nick & "?", vbYesNo + vbQuestion) = vbYes Then
     MsgBox "Silencio tiempo elegido"
    End If
    
Case 36 'Cuenta Regresiva
     CR = InputBox("Cantida de segundos.", "Cuenta regresiva")
     CR2 = InputBox("¿Mapa?, 0 = Mundo.", "Cuenta regresiva")
    If MsgBox("¿Esta seguro que desea ejecutar una cuenta regresiva de " & CR & " segundos?", vbYesNo + vbQuestion) = vbYes Then
        Call WriteCuentaRegresiva(CR, CR2)
    End If
    
Case 39 'part
    tmp = InputBox("¿Particula a colocar?", vbNullString)
    
    If IsNumeric(tmp) Then
        If tmp > 0 And tmp <= 131 Then Call WriteParticulaUsuario(Nick, tmp)
    End If
Case 40 'Oro
     CR = InputBox("Multiplicar: máximo 128.", "Multiplicar")
     CR2 = InputBox("¿Tiempo?: máximo 60.", "Tiempo")
    If MsgBox("¿Esta seguro que desea ingresar un multiplicador de oro x" & CR & ", por " & CR2 & " minutos?", vbYesNo + vbQuestion) = vbYes Then
        Call WriteCuentaRegresiva(CR, CR2)
    End If
Case 41 'Exp
     CR = InputBox("Multiplicar: máximo 128.", "Multiplicar")
     CR2 = InputBox("¿Tiempo?: máximo 60.", "Tiempo")
    If MsgBox("¿Esta seguro que desea ingresar un multiplicador de experiencia x" & CR & ", por " & CR2 & " minutos?", vbYesNo + vbQuestion) = vbYes Then
        Call WriteCuentaRegresiva(CR, CR2)
    End If

Case 42 'Info
    MsgBox "info donador"
Case 43
    MsgBox "haer donad"
    
Case 44
    MsgBox "otorgar cred"
    
Case 45
    MsgBox "reload ban ip"
    
Case 46
    MsgBox "atraer npc"
    
Case 47
    MsgBox "rreset inventario"
    
Case 48
    MsgBox "crear npc"
    
Case 49
    MsgBox "kill npc"
    
    
Case 50
    MsgBox "kill area"
    
Case 51
    MsgBox "ban clan"
    
    
Case 52
    MsgBox "unban ip"
    
    
Case 53
    MsgBox "lastip"
    
    
Case 54
    MsgBox "crear tp"

Case 55
    MsgBox "destp t"
    
Case 56
    MsgBox "map info pk"
    
Case 57
    MsgBox "iembros clan"
     
Case 58
    MsgBox "crear obj"
       
Case 59
    MsgBox "desteuir obj"
     
     
Case 60
    MsgBox "destruir obj area"
     

Case 61
    MsgBox "rajar miembro clan"
    
Case 62
    MsgBox "apass"

Case 63
    MsgBox "showname"







Case 64
    MsgBox "IRCERCA"
 
Case 65
    MsgBox "TELEP   "

Case 67
    MsgBox "INVISIBLE  "

Case 68
    MsgBox "RMATA  "

Case 69
    MsgBox "mod yo"

Case 70
    MsgBox "STAT "

Case 71
    MsgBox "BAL   "

Case 72
    MsgBox "ONLINEGM   "

Case 73
    MsgBox "ONCLAN  "

Case 74
    MsgBox "noestupido74"

Case 75
    MsgBox "CC "

Case 76
    MsgBox "trigger  "

Case 77
    MsgBox "RACC "

Case 78
    MsgBox "NAVE "

Case 79
    MsgBox "GUARDAMAPA "



End Select

reason = vbNullString
Nick = vbNullString
End Sub

  Private Sub cmdActualiza_Click()
    Call WriteRequestUserList
    Call FlushBuffer

End Sub

Private Sub cmdOnlineClan_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub cmdTarget_Click()
    If Check1.Value = 1 Then
     If Text1.Text = vbNullString Then Exit Sub
    Call WriteResponderGm(Text1.Text, txtMsg.Text, "Usuario")
    Else
    Call WriteResponderGm(cboListaUsus.Text, txtMsg.Text, "Usuario")
    End If
End Sub
Private Sub Command1_Click()
Call WriteSystemMessage(txtMsg.Text)
End Sub
 

Private Sub Command4_Click()
MsgBox "RMSG"
End Sub

Private Sub mnuAccion_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuAdmin_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuBan_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub
Private Sub mnuCarcel_Click(Index As Integer)
If Index = 60 Then
    Call cmdAccion_Click(11)
    MsgBox "carcel 60"
    Exit Sub
End If
MsgBox "cartel xx"
'Call Clienttcp.ParseUserCommand("/CARCEL " & cboListaUsus.Text & "@" & "Mal comportamiento" & "@" & Index)
End Sub

Private Sub MnuChange_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuCre_Click(Index As Integer)
If Index = 60 Then
    Call cmdAccion_Click(35)
    MsgBox "cuenta regresiva manual"
    Exit Sub
End If

MsgBox "cuenta regresiva selecc"

End Sub

Private Sub mnuHerramientas_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuIP_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuIrCerca_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuMulti_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuNPC_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuNuevos_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuObject_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuParticula_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuPrem_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuPwChars_Click()
MsgBox "pww chard"
End Sub

Private Sub mnuReload_Click(Index As Integer)
Select Case Index
    Case 1 'Reload objetos
        MsgBox 1
    Case 2 'Reload hechi
        MsgBox 2
    Case 3 'Reload npcs
        MsgBox 3
    Case 4 'Reload sock
        If MsgBox("Al realizar esta acción reiniciará la API de Winsock. Se cerrarán todas las conexiónes.", vbYesNo, "Advertencia") = vbYes Then _
            MsgBox 7
    Case 5 'untnual serve
        MsgBox "puntnual serverini"
    Case 6
    MsgBox " echar todos pjs"
End Select
End Sub

Private Sub mnuSilencio_Click(Index As Integer)
If Index = 60 Then
    Call cmdAccion_Click(35)
    Exit Sub
End If

MsgBox "Silencio tiempo X"

End Sub

Private Sub mnuTelep_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

