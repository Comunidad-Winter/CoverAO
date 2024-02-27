Attribute VB_Name = "modMapas"
Option Explicit
Public Type tRestriccion
    NivelMinimo As Byte
    NivelMaximo As Byte
End Type

Public Type MapBlock

    Blocked As Byte
    Graphic(1 To 4) As Integer
    UserIndex As Integer
    NpcIndex As Integer

    ObjEsFijo As Byte


    
    BlockEsFijo As Byte
 
End Type

'Info del mapa
Type MapInfo

    NumUsers As Integer
    Music As String
    Name As String

    MapVersion As Integer
    Seguro As Byte
    Pk As Boolean
    MagiaSinEfecto As Byte
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
   
    terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    RoboNpcsPermitido As Byte
    
    battle_mode As Byte
    
    RestriccionPorNivel As Byte
    Restriccion As tRestriccion
    
    ChatActivado As Boolean
    
    CaenItems As Boolean
    
    puedeatacar As Boolean
    
    Atacada As Boolean
End Type
 
 
Private Type tMapHeader

    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long

End Type
 
Private Type tDatosBloqueados

    X As Integer
    Y As Integer

End Type
 
Private Type tDatosGrh

    X As Integer
    Y As Integer
    GrhIndex As Long

End Type
 
Private Type tDatosTrigger

    X As Integer
    Y As Integer
    Trigger As Integer

End Type
 
Private Type tDatosLuces

    X As Integer
    Y As Integer
    color As Long
    Rango As Byte

End Type
 
Private Type tDatosParticulas

    X As Integer
    Y As Integer
    Particula As Long

End Type
 
Private Type tDatosNPC

    X As Integer
    Y As Integer
    NpcIndex As Integer

End Type
 
Private Type tDatosObjs

    X As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer

End Type
 
Private Type tDatosTE

    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer

End Type
 
Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type
 
Private Type tMapDat
    map_name As String * 64
    battle_mode As Byte
    backup_mode As Byte
    restrict_mode As String * 4
    music_number As String * 16
    zone As String * 16
    terrain As String * 16
    ambient As String * 16
    base_light As Long
    letter_grh As Long
    extra1 As Long
    extra2 As Long
    extra3 As String * 32
    
End Type

Private MapSize As tMapSize
Private MapDat As tMapDat

Public MapData()                          As MapBlock
Public MapInfo()                          As MapInfo

Public Sub LoadMap(ByVal Map As Long, ByVal MAPFl As String)

    On Error GoTo errh
 
 
 
    Dim fh           As Integer
    Dim MH           As tMapHeader
    Dim Blqs()       As tDatosBloqueados
    Dim L1()         As Long
    Dim L2()         As tDatosGrh
    Dim L3()         As tDatosGrh
    Dim L4()         As tDatosGrh
    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    Dim NPCs()       As tDatosNPC
    Dim TEs()        As tDatosTE
    Dim MapSize      As tMapSize
    Dim MapDat       As tMapDat
 
    Dim i            As Long
    Dim j            As Long
 
 
108    fh = FreeFile
    
110    Dim fTxt As Integer
112    fTxt = FreeFile
    
114    Open MAPFl & ".csm" For Binary Access Read As fh

         
116    Get #fh, , MH
118    Get #fh, , MapSize
120    Get #fh, , MapDat
 
122    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
       
124    Get #fh, , L1
       
126    With MH

128        If .NumeroBloqueados > 0 Then
130            ReDim Blqs(1 To .NumeroBloqueados)
132            Get #fh, , Blqs

 
134        End If
           
136        If .NumeroLayers(2) > 0 Then
138            ReDim L2(1 To .NumeroLayers(2))
140            Get #fh, , L2

142        End If
           
144        If .NumeroLayers(3) > 0 Then
146            ReDim L3(1 To .NumeroLayers(3))
148            Get #fh, , L3

150        End If
           
152        If .NumeroLayers(4) > 0 Then
154            ReDim L4(1 To .NumeroLayers(4))
156            Get #fh, , L4

158        End If

           
160        If .NumeroTriggers > 0 Then
162            ReDim Triggers(1 To .NumeroTriggers)
164            Get #fh, , Triggers


            
166        End If
           
168        If .NumeroParticulas > 0 Then

               
170            ReDim Particulas(1 To .NumeroParticulas)
172            Get #fh, , Particulas

174        End If
            
           
176        If .NumeroLuces > 0 Then
178            ReDim Luces(1 To .NumeroLuces)
180            Get #fh, , Luces

182        End If
           
184        If .NumeroOBJs > 0 Then
186            ReDim Objetos(1 To .NumeroOBJs)
188            Get #fh, , Objetos

             
190        End If
               
192        If .NumeroNPCs > 0 Then
194            ReDim NPCs(1 To .NumeroNPCs)
196            Get #fh, , NPCs


            
198        End If
               
200        If .NumeroTE > 0 Then
202            ReDim TEs(1 To .NumeroTE)
204            Get #fh, , TEs

206        End If
           
208    End With
   
210    Close fh
 
  '  MapInfo(Map).battle_mode = Trim$(MapDat.battle_mod    MapInfo(Map).terreno = Trim$(MapDat.terrain)
Dim terreno As String
 terreno = Trim$(MapDat.terrain)

 
    If terreno = "NOH" Then
        Kill map_path & "Mapa" & Map & ".csm"
    End If
    
    Exit Sub
 
errh:
    Debug.Print "error en mapa numero " & Map

End Sub


Function FileExist(ByVal File As String, _
                   Optional FileType As VbFileAttribute = vbNormal) As Boolean
    '*****************************************************************
    'Se fija si existe el archivo
    '*****************************************************************

    FileExist = LenB(Dir$(File, FileType)) <> 0

End Function
Public Function CargarMapa()

            
    Dim f As Integer
    Dim i As Integer
    Dim j As Integer
    Dim tmpStr As String
    
    f = FreeFile
 
    Dim pompa() As String
    Dim MAX As Integer
    
    i = 0
 
    'Listamos el contenido de la carpeta Dats
    Dim sFilename As String
    
    sFilename = Dir$(map_path)
    
    Do While sFilename > vbNullString
    
     
      sFilename = Dir$()
        i = i + 1
        
    Loop

    
    For j = 1 To i
        Call LoadMap(j, map_path & "Mapa" & j)
    
    Next j
    DoEvents
    
    Exit Function
 
End Function

