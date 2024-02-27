VERSION 5.00
Begin VB.Form frmConversor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conversor de recursos CoverAO"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4605
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
   ScaleHeight     =   1920
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmConversor.frx":0000
      Left            =   960
      List            =   "frmConversor.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convertir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Convertir .dat a locale_es.ind"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   2970
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2940
   End
End
Attribute VB_Name = "frmConversor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Cargando = True Then Exit Sub
Label2.Caption = "Trabajaras con " & Combo1.List(Combo1.ListIndex)
If Combo1.ListIndex >= 0 Then
LoadData (Combo1.ListIndex)
DoEvents
End If
frmConversor.Label1.Caption = "Convertir " & Combo1.List(Combo1.ListIndex) & ".dat" & " a " & "locale_" & Combo1.List(Combo1.ListIndex) & "_es"
End Sub

Private Sub Command1_Click()

If Combo1.ListIndex >= 0 Then
Me.Command1.Enabled = False
Convertir (Combo1.ListIndex)
DoEvents
Me.Command1.Enabled = True
End If
End Sub
Public Function Convertir(ByVal Modo As Byte)

Select Case Modo
Case 0 'Hechizos
Call ConvertirHechizos
Case 1 'Obj
Call ConvertirOBJ
Case 2 'NPCs
Call ConvertisNPCs
Case Else
MsgBox "Error de index"

End Select

End Function
Public Sub ConvertirOBJ()
        Dim i As Integer
        
        Set fs = CreateObject("Scripting.FileSystemObject")

        Set f = fs.GetFile(App.Path & "\conversores\init\locale_obj_es.ind")
        Set fw = f.OpenAsTextStream(8, -2)
        Cargamos = True
        For i = 1 To NumMaximoData
        Debug.Print "Guardando locale_obj_es... " & Round(i / NumMaximoData * 100, 2) & "%"
        fw.WriteLine (OBJData(i).Name & "|" & OBJData(i).desc & "|" & OBJData(i).GrhIndex & "|" & OBJData(i).tipe & "|" & OBJData(i).MaxDef & "|" & OBJData(i).MinDef & "|" & OBJData(i).MaxHit & "|" & OBJData(i).MinHit & "|" & OBJData(i).CreaLuz & "|" & OBJData(i).RangoLuz & "|" & OBJData(i).Snd1 & "|" & OBJData(i).Snd2 & "|" & OBJData(i).Snd3 & "|" & OBJData(i).Nivel & "|" & OBJData(i).CreaParticulaPiso)
        Next i
        fw.Close
        
        Label2.Caption = "Conversión exitosa"
End Sub
Public Sub ConvertirHechizos()
        Dim i As Integer
        
        Set fs = CreateObject("Scripting.FileSystemObject")

        Set f = fs.GetFile(App.Path & "\conversores\init\locale_spl_es.ind")
        Set fw = f.OpenAsTextStream(8, -2)
        Cargamos = True
        For i = 1 To NumMaximoData
        Debug.Print "Guardando locale_spl_es... " & Round(i / NumMaximoData * 100, 2) & "%"
        fw.WriteLine (SpellData(i).Name & "|" & SpellData(i).desc & "|" & SpellData(i).strHechizeroMsg & "|" & SpellData(i).strTargetMsg & "|" & SpellData(i).strOwnMsg & "|" & SpellData(i).strMagicas & "|" & SpellData(i).strTarget & "|" & SpellData(i).ManaRequerido & "|" & SpellData(i).StaRequerido & "|" & SpellData(i).SkillRequerido)
        Next i
        fw.Close
        
        Label2.Caption = "Conversión exitosa"
        
End Sub

Public Sub ConvertisNPCs()
        Dim i As Integer
        
        Set fs = CreateObject("Scripting.FileSystemObject")

        Set f = fs.GetFile(App.Path & "\conversores\init\locale_npc_es.ind")
        Set fw = f.OpenAsTextStream(8, -2)
        Cargamos = True
        For i = 1 To NumMaximoData
        Debug.Print "Guardando locale_npc_es... " & Round(i / NumMaximoData * 100, 2) & "%"
        fw.WriteLine (NpcData(i).Name & "|" & NpcData(i).status & "|" & NpcData(i).desc & "|" & NpcData(i).Hostil & "|" & NpcData(i).MinHp & "|" & NpcData(i).MaxHp & "|" & NpcData(i).MinHit & "|" & NpcData(i).MaxHit & "|" & NpcData(i).Nivel & "|" & NpcData(i).DaOro & "|" & NpcData(i).DaExp & "|" & NpcData(i).Defensa & "|" & NpcData(i).PoderEvasion & "|" & NpcData(i).PoderAtaque & "|" & NpcData(i).Head & "|" & NpcData(i).Body & "|" & NpcData(i).CascoAnim & "|" & NpcData(i).ShieldAnim & "|" & NpcData(i).WeaponAnim & "|" & NpcData(i).heading & "|" & NpcData(i).stats & "|" & NpcData(i).RandomDrop)
        
        Next i
        fw.Close
        
        Label2.Caption = "Conversión exitosa"
End Sub

