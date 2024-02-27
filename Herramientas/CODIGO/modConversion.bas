Attribute VB_Name = "modConversion"

Option Explicit
Dim ConversorPath As String
Dim InitPaths As String
Public Type SpellData
    Name               As String 'Nombre del obj
    desc              As String
    strHechizeroMsg As String
    strTargetMsg As String
    strOwnMsg As String
    strMagicas As String
    strTarget As String
    SkillRequerido As Integer
    ManaRequerido As Integer
    StaRequerido As Integer
End Type

Public SpellData()                          As SpellData

Public Type NpcData
    Name As String 'Nombre del obj
    
    Nivel As Integer 'nivel
    
    DaOro As Long
    DaExp As Long
    
    MinHit As Long
    MaxHit As Long
    
    Defensa As Integer
    
    MinHp As Long
    MaxHp As Long
    
    PoderAtaque As Integer
    PoderEvasion As Integer
    
    Head As Integer
    Body As Integer
    CascoAnim As Integer
    ShieldAnim As Integer
    WeaponAnim As Integer
    heading As Byte

    status  As String
    desc As String
    Hostil As Byte

    stats As Integer
    RandomDrop As String
End Type
 
            
Public NpcData()                          As NpcData


Public Type OBJData
    Name               As String 'Nombre del obj
    desc      As String
    tipe As Byte
    GrhIndex As Integer
    MaxHit As Integer
    MinHit As Integer
    MaxDef As Integer
    MinDef As Integer
    CreaLuz As String
    RangoLuz As String
    Snd1 As Integer 'Snd Equipar
    Snd2 As Integer 'Snd Golpe
    Snd3 As Integer 'Snd fallas
    Nivel As Integer
    CreaParticulaPiso As Integer
End Type

Public OBJData()                          As OBJData

Public NumMaximoData                        As Integer
Public Cargando As Boolean
Public Sub LoadData(ByVal Index As Byte)
    
    On Error GoTo Errhandler
    ConversorPath = App.Path & "\Conversores\Dat\"
    InitPaths = App.Path & "\Conversores\Init\"
    
    Select Case Index
    
    Case 0
    Call LoadHechizosDat
    Case 1
    Call LoadOBJDat
    Case 2
    Call LoadNPCsDat
    
    Case Else
    MsgBox "Error al seleccionar index"
    Exit Sub
    End Select
 
    Exit Sub
 
Errhandler:
    MsgBox "error cargando el index:  " & Index
    
End Sub
Public Sub LoadHechizosDat()
    On Error Resume Next
 
    If Cargando Then Exit Sub
    
    Cargando = True
    Dim Object As Integer
    Dim Leer   As New clsIniReader
    Set Leer = New clsIniReader
 
    
    Call Leer.Initialize(ConversorPath & "hechizos.dat")
     
     
    NumMaximoData = Val(Leer.GetValue("INIT", "NumeroHechizos"))
   
    ReDim Preserve SpellData(1 To NumMaximoData) As SpellData
        
        For Object = 1 To NumMaximoData
        frmConversor.Label2.Caption = "Cargando Hechizos... " & Round(Object / NumMaximoData * 100, 2) & "%"
        Debug.Print "Cargando Hechizos.dat... " & Round(Object / NumMaximoData * 100, 2) & "%"
        SpellData(Object).Name = Leer.GetValue("HECHIZO" & Object, "Nombre") 'Curar Veneno
        SpellData(Object).desc = Leer.GetValue("HECHIZO" & Object, "Desc") 'Cura el envenenamiento
        SpellData(Object).strHechizeroMsg = Leer.GetValue("HECHIZO" & Object, "HechizeroMsg") 'Has curado a
        SpellData(Object).strTargetMsg = Leer.GetValue("HECHIZO" & Object, "TargetMsg")  ' '""" Te ha curado el envenenamiento
        SpellData(Object).strOwnMsg = Leer.GetValue("HECHIZO" & Object, "PropioMsg") ''Te has curado
        SpellData(Object).strMagicas = Leer.GetValue("HECHIZO" & Object, "PalabrasMagicas") ''Nihil Ved
        SpellData(Object).strTarget = Leer.GetValue("HECHIZO" & Object, "Target")
        SpellData(Object).SkillRequerido = Leer.GetValue("HECHIZO" & Object, "MinSkill")
        SpellData(Object).ManaRequerido = Leer.GetValue("HECHIZO" & Object, "ManaRequerido")
        SpellData(Object).StaRequerido = Leer.GetValue("HECHIZO" & Object, "StaRequerido")
        DoEvents
        Next Object
        
        
    Set Leer = Nothing
    Cargando = False
    frmConversor.Label2 = "Finalizada carga Hechizos.dat."
    frmConversor.Command1.Enabled = True
    Exit Sub
 
Errhandler:
    MsgBox "error cargando Hechizos " & Err.Number & ": " & Err.Description
    
End Sub



Public Sub LoadOBJDat()
    On Error Resume Next
 
    If Cargando Then Exit Sub
    
    Cargando = True
    Dim Object As Integer
    Dim LuzIndex() As String
    Dim Leer   As New clsIniReader
    Set Leer = New clsIniReader
 
    
    Call Leer.Initialize(ConversorPath & "obj.dat")
     
     
    NumMaximoData = Val(Leer.GetValue("INIT", "NumOBJs"))
   
    ReDim Preserve OBJData(1 To NumMaximoData) As OBJData
        
        For Object = 1 To NumMaximoData
        frmConversor.Label2.Caption = "Cargando obj.dat... " & Round(Object / NumMaximoData * 100, 2) & "%"
        Debug.Print "Cargando obj.dat... " & Round(Object / NumMaximoData * 100, 2) & "%"
        OBJData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")
        OBJData(Object).desc = Leer.GetValue("OBJ" & Object, "Texto")
        OBJData(Object).GrhIndex = Leer.GetValue("OBJ" & Object, "GrhIndex")
        OBJData(Object).tipe = Leer.GetValue("OBJ" & Object, "ObjType")
        OBJData(Object).MaxDef = Leer.GetValue("OBJ" & Object, "MAXDEF")
        OBJData(Object).MinDef = Leer.GetValue("OBJ" & Object, "MINDEF")
        OBJData(Object).MaxHit = Leer.GetValue("OBJ" & Object, "MAXHIT")
        OBJData(Object).MinHit = Leer.GetValue("OBJ" & Object, "MINHIT")
        
        LuzIndex() = Split(Leer.GetValue("OBJ" & Object, "CreaLuz"), ":")
        OBJData(Object).CreaLuz = (LuzIndex(1))
        OBJData(Object).RangoLuz = (LuzIndex(0))
        
        OBJData(Object).Snd1 = Leer.GetValue("OBJ" & Object, "Snd1")
        OBJData(Object).Snd2 = Leer.GetValue("OBJ" & Object, "Snd2")
        OBJData(Object).Snd3 = Leer.GetValue("OBJ" & Object, "Snd3")
        OBJData(Object).Nivel = Leer.GetValue("OBJ" & Object, "MinELV")
         OBJData(Object).CreaParticulaPiso = Leer.GetValue("OBJ" & Object, "CreaParticulaPiso")
          
        DoEvents
        
        Next Object
        
    Set Leer = Nothing
    Cargando = False
    frmConversor.Label2 = "Finalizada carga obj.dat."
    frmConversor.Command1.Enabled = True
    Exit Sub
 
Errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description
    
End Sub



Public Sub LoadNPCsDat()
    On Error Resume Next
 
    If Cargando Then Exit Sub
    
    Cargando = True
    Dim Object As Integer
    Dim Leer   As New clsIniReader
    Set Leer = New clsIniReader
 
    
    Call Leer.Initialize(ConversorPath & "npcs.dat")
     
     
    NumMaximoData = Val(Leer.GetValue("INIT", "NumNPCs"))
   
    ReDim Preserve NpcData(1 To NumMaximoData) As NpcData
        
        For Object = 1 To NumMaximoData

        frmConversor.Label2.Caption = "Cargando npcs.dat... " & Round(Object / NumMaximoData * 100, 2) & "%"
        Debug.Print "Cargando npcs.dat... " & Round(Object / NumMaximoData * 100, 2) & "%"
        NpcData(Object).Name = CStr(Leer.GetValue("NPC" & Object, "Name"))
        NpcData(Object).MaxHp = Val(Leer.GetValue("NPC" & Object, "MaxHP"))
        NpcData(Object).MinHp = Val(Leer.GetValue("NPC" & Object, "MinHP"))
        NpcData(Object).Hostil = Val(Leer.GetValue("NPC" & Object, "Hostile"))
        NpcData(Object).desc = CStr(Leer.GetValue("NPC" & Object, "Desc"))
        NpcData(Object).status = CStr(Leer.GetValue("NPC" & Object, "Name2"))
        
        If Not Len(NpcData(Object).status) > 2 Then
        NpcData(Object).status = Leer.GetValue("NPC" & Object, "Status")
            Select Case NpcData(Object).status
                Case 1 'Imperial
                NpcData(Object).status = "Sagrada Orden"
                Case 2 'Republica
                NpcData(Object).status = "Milicia Republicana"
                Case 4 'Caos
                NpcData(Object).status = "Fuerzas del Caos"
                Case 5 'Gm
                NpcData(Object).status = "CoverAO Staff"
                Case 6 'Sagra
                NpcData(Object).status = "Sagrada Orden"
            End Select
        End If
        
        NpcData(Object).stats = Val(Leer.GetValue("NPC" & Object, "Status"))
        
        If NpcData(Object).stats = 1 Then
                NpcData(Object).stats = 2
        ElseIf NpcData(Object).stats = 2 Then
                NpcData(Object).stats = 3
        
        ElseIf NpcData(Object).stats = 3 Then
            NpcData(Object).stats = 1
        ElseIf NpcData(Object).stats = 4 Then
            NpcData(Object).stats = 4
        ElseIf NpcData(Object).stats = 5 Then
            NpcData(Object).stats = 5
        ElseIf NpcData(Object).stats = 6 Then
            NpcData(Object).stats = 6
        Else
        
            NpcData(Object).stats = 1
        End If
        
        
        NpcData(Object).MinHit = Val(Leer.GetValue("NPC" & Object, "MinHit"))
        NpcData(Object).MaxHit = Val(Leer.GetValue("NPC" & Object, "MaxHit"))
        
        NpcData(Object).Nivel = Val(Leer.GetValue("NPC" & Object, "Nivel"))
        
        If NpcData(Object).Nivel <= 0 Then
            NpcData(Object).Nivel = 100
        End If
        
        NpcData(Object).DaOro = Val(Leer.GetValue("NPC" & Object, "GiveGLD"))
        NpcData(Object).DaExp = Val(Leer.GetValue("NPC" & Object, "GiveEXP"))
                
        NpcData(Object).Defensa = Val(Leer.GetValue("NPC" & Object, "DEF"))
        NpcData(Object).PoderAtaque = Val(Leer.GetValue("NPC" & Object, "PoderAtaque"))
        NpcData(Object).PoderEvasion = Val(Leer.GetValue("NPC" & Object, "PoderEvasion"))
        
        NpcData(Object).Head = Val(Leer.GetValue("NPC" & Object, "Head"))
        NpcData(Object).Body = Val(Leer.GetValue("NPC" & Object, "Body"))

        NpcData(Object).CascoAnim = Val(Leer.GetValue("NPC" & Object, "CascoAnim"))
        NpcData(Object).ShieldAnim = Val(Leer.GetValue("NPC" & Object, "ShieldAnim"))
        
        NpcData(Object).WeaponAnim = Val(Leer.GetValue("NPC" & Object, "WeaponAnim"))
        NpcData(Object).heading = Val(Leer.GetValue("NPC" & Object, "heading"))
        
        If Leer.GetValue("NPC" & Object, "RandomDrop") <> "" Then
            NpcData(Object).RandomDrop = Leer.GetValue("NPC" & Object, "RandomDrop")
        Else
            NpcData(Object).RandomDrop = 0
        End If
        
        DoEvents
        
        Next Object
        
    Set Leer = Nothing
    Cargando = False
    frmConversor.Label2 = "Finalizada carga npcs.dat."
    frmConversor.Command1.Enabled = True
    Exit Sub
 
Errhandler:
    MsgBox "error cargando npcs " & Err.Number & ": " & Err.Description
    
End Sub
