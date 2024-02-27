VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
Public Sub ParseUserCommand(ByVal strCmd As String)

    On Error GoTo errorhandler
    
1    Dim strArg As String, strCompa As String
2    Dim bFound As Boolean

3    If Len(strCmd) > 255 Then strCmd = mid$(strCmd, 1, 255)

4    bFound = True
5    strCompa = UCase$(strCmd)

6    Select Case mid$(strCompa, 2)
 
        Case Locale_CMD_Get(1) '/PING
            Call AddtoRichTextBox("Ping: " & MSRender & "ms", 0, 0, 0, 0, 0, True)
        
        Case Locale_CMD_Get(2) '/SALIR
            Call WriteQuit
        
        Case Locale_CMD_Get(3) '/MEDITAR
            If DeadCheck Then Exit Sub
            Call WriteMeditate
        
        Case Locale_CMD_Get(4) '/ONLINE
            Call WriteOnline

        Case Locale_CMD_Get(5) '/ONLINECLAN
            Call WriteGuildOnline
    
        Case Locale_CMD_Get(6), Locale_CMD_Get(7), Locale_CMD_Get(8) '/GM, /DM, /AYUDA
            frmGMAyuda.Show vbModeless, frmMain
            frmGMAyuda.TxtSoporte.SetFocus

        Case Locale_CMD_Get(9) '/BALANCE
            If DeadCheck Then Exit Sub
            MsgBox "balance"
    
        Case Locale_CMD_Get(10) '/EST
            Call WriteRequestStats
    
        Case Locale_CMD_Get(11) '/BOVEDA
            If DeadCheck Then Exit Sub
            Call WriteBankStart
    
        Case Locale_CMD_Get(12) '/COMERCIAR
            If DeadCheck Then Exit Sub
            If Not Comerciando Then Call WriteCommerceStart
        
        Case Locale_CMD_Get(13) '/INFORMACION
            Call WriteInformation
        
        Case Locale_CMD_Get(14) '/ENLISTAR
            If DeadCheck Then Exit Sub
            Call WriteEnlist
        
        Case Locale_CMD_Get(16) '/FUNDARCLAN
            If DeadCheck Then Exit Sub
            Call WriteGuildFundate
        
        Case Locale_CMD_Get(17) '/SALIRCLAN
            If DeadCheck Then Exit Sub
            Call WriteGuildLeave
        
        Case Locale_CMD_Get(18) '/CERRARCLAN
            If DeadCheck Then Exit Sub
            Call writeCloseGuild
     
        Case Locale_CMD_Get(20) '/DIVORCIO
            If DeadCheck Then Exit Sub
            Call writeDivorciar
        
        Case Locale_CMD_Get(21) '/PENA
            MsgBox "Pena"
        
        Case Locale_CMD_Get(22) '/HOGAR
            Call WriteSeleccionarHogar(0)
        
        Case Locale_CMD_Get(28) '/RETIRAR
            If DeadCheck Then Exit Sub
            frmPregunta.SetAccion 2
            
            If frmMain.Visible Then frmPregunta.Show , frmMain
            'Call WriteRetirarFaccion
            
        Case Locale_CMD_Get(38) '/FLIST
            Call WriteOnAmigoandMapa
        
        Case Locale_CMD_Get(39) '/PERDON
            If DeadCheck Then Exit Sub
            MsgBox "Perdon"
        
        Case Locale_CMD_Get(41) '/EVENTOS
            Call WriteHayEventos
        
        Case Locale_CMD_Get(42) '/SKIN
            If DeadCheck Then Exit Sub
            Call WriteAbrirForms(1)
        
        Case Else
        
            bFound = False
    End Select
 
7    If bFound Then Exit Sub

8    bFound = True

9    strCompa = UCase$(General_Field_Read(1, strCmd, " "))
 
10    Select Case mid$(strCompa, 2)
    
        Case Locale_CMD_Get(15) '/DUELO
            If DeadCheck Then Exit Sub
            MsgBox "Duelo"
        
        Case Locale_CMD_Get(19) '/PROPONER
        
            If DeadCheck Then Exit Sub
            
            strArg = Trim(Right$(strCmd, Len(strCmd) - 9))
             
            If Len(strArg) > 0 Then
                Call WriteCasamiento(strArg, 0)
            End If

        Case Locale_CMD_Get(23) '/GMSG
            CurrentUser.SendingType = 5
            Call WriteGMMessage(Right$(strCmd, Len(strCmd) - 6))
        
        Case Locale_CMD_Get(24) '/CMSG
            If DeadCheck Then Exit Sub
            CurrentUser.SendingType = 4
            Call WriteGuildMessage(Right$(strCmd, Len(strCmd) - 6))
        
        Case Locale_CMD_Get(25) '/GRMG
            CurrentUser.SendingType = 6
            MsgBox Right$(strCmd, Len(strCmd) - 6)
                
        Case Locale_CMD_Get(26) '/FMMG
            CurrentUser.SendingType = 6
            MsgBox Right$(strCmd, Len(strCmd) - 6)
            
        Case Locale_CMD_Get(27) '/IGNORAR
            strArg = Right$(strCmd, Len(strCmd) - 9)
            If strArg <> "" Then Call frmOpciones.AgregarIgnorado(strArg)
            
        Case Locale_CMD_Get(28) '/RETIRAR
            If DeadCheck Then Exit Sub
            strArg = Right$(strCmd, Len(strCmd) - 9)
            MsgBox strArg
            'If ValidNumber(argumentos(1), eNumber_Types.ent_Long) Then
            
            'Call WriteBankExtractGold(strArg)
    
        Case Locale_CMD_Get(29) '/DEPOSITAR
            If DeadCheck Then Exit Sub
            strArg = Right$(strCmd, Len(strCmd) - 11)
            Call WriteBankDepositGold(strArg)
            
        Case Locale_CMD_Get(30) '/APOSTAR
            If DeadCheck Then Exit Sub
            strArg = Right$(strCmd, Len(strCmd) - 9)
            Call WriteGamble(strArg)
            
        Case Locale_CMD_Get(32) '/CENTINELA
            strArg = Right$(strCmd, Len(strCmd) - 11)
            Call WriteCentinelReport(CInt(strArg))
            
        Case Locale_CMD_Get(33) '/DESC
            If DeadCheck Then Exit Sub
            strArg = Right$(strCmd, Len(strCmd) - 6)
            Call WriteChangeDescription(strArg)
            
        Case Locale_CMD_Get(34) '/FADD
            If DeadCheck Then Exit Sub
            Call WriteAddAmigo(Right$(strCmd, Len(strCmd) - 6), 1)
            
        Case Locale_CMD_Get(35) '/FDEL
            If DeadCheck Then Exit Sub
            Call WriteDelAmigo(Right$(strCmd, Len(strCmd) - 6))
            
        Case Locale_CMD_Get(36) '/FMSG
            Call WriteMsgAmigo(Right$(strCmd, Len(strCmd) - 6))
            
        Case Locale_CMD_Get(37) '/FACCEPT
            If DeadCheck Then Exit Sub
            Call WriteAddAmigo(Right$(strCmd, Len(strCmd) - 6), 2)
                
        Case Locale_CMD_Get(40) '/VOTO
            If DeadCheck Then Exit Sub
            Call WriteGuildVote(Right$(strCmd, Len(strCmd) - 6))
            
        Case Locale_CMD_Get(43) ' /DESCONECTAR
                strArg = (Right$(strCmd, Len(strCmd) - 13))
                If Len(strArg) > 0 Then Call WriteDesconectarCuenta(strArg)
            
        Case Locale_CMD_Get(44) ' /DONADOR
             strArg = Trim(Right$(strCmd, Len(strCmd) - 9))
             If Len(strArg) > 0 Then Call WriteDonador(strArg)
             
        Case Else
        
        CurrentUser.SendingType = 1
        bFound = False
        
    End Select


    '''''''Descomentar cuando se pongan comandos GM sin argumento
    'If bFound Then Exit Sub
    'bFound = True
    'strCompa = UCase$(strCmd)
    
    'Select Case mid$(strCompa, 2) 'GMM Commands Solo sin argumento
    '    Case Else
    '        bFound = False
    'End Select

11    If bFound Then Exit Sub

12    bFound = True
13    strCompa = UCase$(General_Field_Read(1, strCmd, " "))
14    strArg = mid$(strCmd, Len(strCompa) + 2)

15    Select Case mid$(strCompa, 2) 'GMM Commands Solo con argumento
        Case Locale_CMD_Get(31) '/REM
            Call WriteComment(strArg)
        
        Case Else
            Call AddtoRichTextBox("¡Comando no reconocido!", 0, 0, 0, 0, 0, 0, 4)
            
    End Select
    
16    Exit Sub

errorhandler:
     Call RegistrarError(Err.number, Err.Description, "clsClientTCP.ParseUserCommand", Erl)
     Resume Next
End Sub

Public Function DeadCheck() As Boolean

If CurrentUser.Muerto Then
    Call AddtoRichTextBox(Locale_GUI_Frase(245), 0, 0, 0, 0, 0, 0, 12)
    DeadCheck = True
End If

End Function
Public Function ActualizarEst(Optional ByVal MaxHP As Integer = -1, Optional ByVal MinHP As Integer = -1, Optional ByVal MaxMAN As Integer = -1, _
    Optional ByVal MinMAN As Integer = -1, Optional ByVal MaxSTA As Integer = -1, Optional ByVal MinSTA As Integer = -1, _
    Optional ByVal GLD As Long = -1, Optional ByVal Nivel As Integer = -1, Optional ByVal PasarNivel As Long = -1, Optional ByVal EXP As Long = -1, _
    Optional ByVal Fuerza As Integer = -1, Optional ByVal Agilidad As Integer = -1, _
    Optional ByVal MinHAM As Integer = -1, Optional ByVal MaxHAM As Integer = -1, _
    Optional ByVal MinSED As Integer = -1, Optional ByVal MaxSED As Integer = -1, _
    Optional ByVal ActualizarTodos As Boolean = False)

Dim ActualizarCual As Byte

If MaxHP <> -1 Then
    CurrentUser.UserMaxHP = MaxHP
    ActualizarCual = 1
End If

If MinHP <> -1 Then

    If MinHP < 0 Then MinHP = 0
    CurrentUser.UserMinHP = MinHP
    ActualizarCual = 1
    
End If

If MaxMAN <> -1 Then
    CurrentUser.UserMaxMAN = MaxMAN
    ActualizarCual = 2
End If

If MinMAN <> -1 Then
    CurrentUser.UserMinMAN = MinMAN
    
    If CurrentUser.UserMinMAN > CurrentUser.UserMaxMAN Then _
        CurrentUser.UserMinMAN = CurrentUser.UserMaxMAN
        
    ActualizarCual = 2
End If

If MaxSTA <> -1 Then
    CurrentUser.UserMaxSTA = MaxSTA
    ActualizarCual = 3
End If

If MinSTA <> -1 Then

    If CurrentUser.UserMinSTA > CurrentUser.UserMaxSTA Then _
        CurrentUser.UserMinSTA = CurrentUser.UserMaxSTA

    CurrentUser.UserMinSTA = MinSTA
    ActualizarCual = 3
End If

If GLD <> -1 Then
    CurrentUser.UserGLD = GLD
    ActualizarCual = 4
End If

If Nivel <> -1 Then
    CurrentUser.UserLvl = Nivel
    ActualizarCual = 5
End If

If PasarNivel <> -1 Then
    CurrentUser.UserPasarNivel = PasarNivel
    ActualizarCual = 5
End If
    
If EXP <> -1 Then
    CurrentUser.UserExp = EXP
    ActualizarCual = 5
End If

If Fuerza <> -1 Then
    frmMain.lblStrg = Fuerza
    frmMain.lblStrg.ForeColor = IIf(Fuerza < 35, &HFFFFFF, &H40C0&)
End If

If Agilidad <> -1 Then
    frmMain.lblDext = Agilidad
    frmMain.lblDext.ForeColor = IIf(Agilidad < 35, &HFFFFFF, &H40C0&)
End If

If MinHAM <> -1 Then
    CurrentUser.UserMinHAM = MinHAM
    ActualizarCual = 6
End If

If MinSED <> -1 Then
    CurrentUser.UserMinAGU = MinSED
    ActualizarCual = 6
End If

If MaxHAM <> -1 Then
    CurrentUser.UserMaxHAM = MaxHAM
    ActualizarCual = 6
End If

If MaxSED <> -1 Then
    CurrentUser.UserMaxAGU = MaxSED
    ActualizarCual = 6
End If

If Not ActualizarTodos Then
    Select Case ActualizarCual
        Case 1
            Call ActualizarHP
        Case 2
            Call ActualizarMAN
        Case 3
            Call ActualizarSTA
        Case 4
            Call ActualizarGLD
        Case 5
            Call ActualizarExp
        Case 6
            Call ActualizarHAM
    End Select
Else
    Call ActualizarHP
    Call ActualizarMAN
    Call ActualizarSTA
    Call ActualizarGLD
    Call ActualizarExp
    Call ActualizarHAM
End If

End Function
Private Sub ActualizarHP()

If CurrentUser.UserMinHP <= 0 Then
    CurrentUser.UserMinHP = 0
    CurrentUser.Muerto = True
    Engine_Scroll_Pixels scroll_pixels_per_frameBackUp * VelocidadMuerto
    frmMain.lblVida.Caption = CurrentUser.UserMinHP & "/" & CurrentUser.UserMaxHP
    frmMain.shpvida.Width = (((CurrentUser.UserMinHP / 100) / (CurrentUser.UserMaxHP / 100)) * 91)
    frmMain.shpvida.FillColor = &H808080
Else
    If CurrentUser.Muerto Then
        CurrentUser.Muerto = False
    End If
    
    If CurrentUser.Logged Then
        If (CurrentUser.Montando = False) Then
            Engine_Scroll_Pixels scroll_pixels_per_frameBackUp
        End If
    End If
    
    frmMain.lblVida.Caption = CurrentUser.UserMinHP & "/" & CurrentUser.UserMaxHP
    frmMain.shpvida.Width = (((CurrentUser.UserMinHP / 100) / (CurrentUser.UserMaxHP / 100)) * 91)
    frmMain.shpvida.FillColor = &HC0&
End If

End Sub
Private Sub ActualizarGLD()
frmMain.GldLbl.Caption = CurrentUser.UserGLD
End Sub
Private Sub ActualizarMAN()

If CurrentUser.UserMaxMAN > 0 Then
    frmMain.shpmana.Width = (((CurrentUser.UserMinMAN + 1 / 100) / (CurrentUser.UserMaxMAN + 1 / 100)) * 91)
    frmMain.lblMana.Visible = True
    frmMain.lblMana.Caption = CurrentUser.UserMinMAN & "/" & CurrentUser.UserMaxMAN
Else
    frmMain.shpmana.Width = 0
    frmMain.lblMana.Visible = False
End If

End Sub
Private Sub ActualizarSTA()
frmMain.shpEnergia.Width = (((CurrentUser.UserMinSTA / 100) / (CurrentUser.UserMaxSTA / 100)) * 91)
frmMain.lblEnergia.Caption = CurrentUser.UserMinSTA & "/" & CurrentUser.UserMaxSTA
End Sub
Private Sub ActualizarExp()

frmMain.lblLvl.Caption = CurrentUser.UserLvl

Call UserExpPerc

If CurrentUser.UserPercExp <> 0 Then
    frmMain.ExpShp.Width = (((CurrentUser.UserExp / 100) / (CurrentUser.UserPasarNivel / 100)) * 121)
Else
    frmMain.ExpShp.Width = 0
End If
        
frmMain.Label2(1).Caption = IIf(frmMain.UltPos = 1, CurrentUser.UserExp & "/" & CurrentUser.UserPasarNivel, CurrentUser.UserPercExp & "%")

If CurrentUser.UserPasarNivel = 0 Then
    frmMain.Label2(1).Caption = Locale_GUI_Frase(173)
End If

End Sub

Private Sub ActualizarHAM()

frmMain.shpSed.Width = (((CurrentUser.UserMinAGU / 100) / (CurrentUser.UserMaxAGU / 100)) * 91)
frmMain.shpHambre.Width = (((CurrentUser.UserMinHAM / 100) / (CurrentUser.UserMaxHAM / 100)) * 91)
frmMain.lblHambre.Caption = CurrentUser.UserMinHAM & "/" & CurrentUser.UserMaxHAM
frmMain.lblSed.Caption = CurrentUser.UserMinAGU & "/" & CurrentUser.UserMaxAGU

End Sub

