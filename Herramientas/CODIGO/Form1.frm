VERSION 5.00
Begin VB.Form frmDescomprensor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descomprensor/comprensor recursos CoverAO"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Descomprimir"
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   840
      List            =   "Form1.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comprimir"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¿Con que recursos desea trabajar?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   0
      Width           =   2940
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2385
      TabIndex        =   1
      Top             =   2400
      Width           =   75
   End
End
Attribute VB_Name = "frmDescomprensor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Combo1_Click()

Label1.Caption = "Trabajaras con " & Combo1.List(Combo1.ListIndex)
If Combo1.ListIndex = 8 Then
Text1.Enabled = True
Else
Text1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
        

        Dim a As String
        Dim b As Byte, c As Byte
        
        b = Combo1.ListIndex
        strCurSkinName = Text1.Text
        Select Case b
        Case 0
        a = graph_path
        c = Graphics
         
        Case 1
        a = SOUNDS_PATH
        c = Midi
        
        Case 2
        a = MP3_PATH
        c = MP3
        
        Case 3
        a = SOUNDS_PATH
        c = Wav
        
        Case 4
        a = SCRIPT_PATH
        c = Scripts

        Case 5
        a = PATCH_PATH
        c = Patch
        
        Case 6
        a = INTERFACE_PATH
        c = Interface
        Case 7
        Debug.Print Time
        Label1.Caption = "Espera que carguen los mapas"
        a = map_path
        c = Maps
        CargarMapa
        Label1.Caption = "Listo carga de mapas"
        Case 8
 
        a = SKIN_PATH
        c = Skins
        End Select
     
        Label1.Caption = "Comprimiendo " & Combo1.List(Combo1.ListIndex)
        Compress_Files c, a, resource_path
        Label1.Caption = "Finalizado"
End Sub

Private Sub Command2_Click()

        Dim a As String
        Dim b As Byte, c As Byte
        
        b = Combo1.ListIndex

        strCurSkinName = Text1.Text
        Select Case b
        Case 0
        c = Graphics
         
        Case 1
        c = Midi
        
        Case 2
        c = MP3
        
        Case 3
        c = Wav
        
        Case 4
        c = Scripts

        Case 5
        c = Patch
        
        Case 6
        c = Interface
        Case 7
        c = Maps
        
        Case 8
        c = Skins
         
        End Select

        Label1.Caption = "Descomprimiendo " & Combo1.List(Combo1.ListIndex)
        Extract_All_Files c, resource_path
        Label1.Caption = "Finalizado"
End Sub

Private Sub Form_Load()

'Comprensor / Descompensor
resource_path = App.Path & "\Comprensor\Recursos\"
graph_path = App.Path & "\Comprensor\Descomprimido\graficos\"
MP3_PATH = App.Path & "\Comprensor\Descomprimido\MP3\"
SOUNDS_PATH = App.Path & "\Comprensor\Descomprimido\Sounds\"
SCRIPT_PATH = App.Path & "\Comprensor\Descomprimido\Init\"
PATCH_PATH = App.Path & "\Comprensor\Descomprimido\Patches\"
INTERFACE_PATH = App.Path & "\Comprensor\Descomprimido\interface\"
map_path = App.Path & "\Comprensor\Descomprimido\mapas\"
SKIN_PATH = App.Path & "\Comprensor\Descomprimido\skins\"
'Comprensor descomprensor
End Sub

