Attribute VB_Name = "Mod_AutoUsar"
'---------------------------------------------------------------------------------------
' CoverAO - V1.0                                                                       '
' Fecha     : 19/07/2020                                                               '
' Module    : Mod_AutoUsar                                                             '
'---------------------------------------------------------------------------------------

Option Explicit

Public Function Input_Key_Get(ByVal key_code As Byte) As Boolean
'--------------------------------------------------
'Author: Aaron Perkins - Juan Martín Sotuyo Dodero
'Now we use DirectInput Keyboard
'Last Modify Date: 10/07/2002
'Agradecimiento a Ladder
'--------------------------------------------------
 
 Input_Key_Get = (GetKeyState(key_code) < 0)
 
End Function
Public Function esArco(ByVal Obj As Integer) As Boolean

    Select Case Obj
    Case 989, 1355, 1001, 709, 479, 899, 478, 655, 564, 666, 749, 138: esArco = True
    Case Else: esArco = False
    End Select

End Function
Public Function esArrojadiza(ByVal Obj As Integer) As Boolean

    Select Case Obj
    Case 576, 671, 571, 656, 742, 720, 741, 980, 1594, 1241, 1595, 1596: esArrojadiza = True
    Case Else: esArrojadiza = False
    End Select

End Function
Public Function esHerramienta(ByVal Obj As Integer)

    Select Case Obj
    
    Case 881, 138 'Caña, red
        Call AddtoRichTextBox(General_Locale_SMG(462, 0), 0, 0, 0, 0, 0, 0, 12)
        
    Case 187 'Piquete
        Call AddtoRichTextBox(General_Locale_SMG(463, 0), 0, 0, 0, 0, 0, 0, 12)
        
    Case 127, 885 'Tijeras, hacha
        Call AddtoRichTextBox(General_Locale_SMG(464, 0), 0, 0, 0, 0, 0, 0, 12)
          
    Case 192, 193, 194 'Minerales
        Call AddtoRichTextBox(General_Locale_SMG(465, 0), 0, 0, 0, 0, 0, 0, 12)
          
    Case 386, 387, 388 'Lingos
        Call AddtoRichTextBox(General_Locale_SMG(467, 0), 0, 0, 0, 0, 0, 0, 12)
        
    Case 389 'Martillo
        Call AddtoRichTextBox(General_Locale_SMG(466, 0), 0, 0, 0, 0, 0, 0, 12)
        
    Case Else
        Call WriteUseItem(Inventario.SelectedItem)
        
    End Select
End Function
Public Sub UsarItem()
     
    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub

 
     If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then

        ' Hacemos acción del doble clic correspondiente
        Dim ObjType As Integer
    
        ObjType = Inventario.ObjType(Inventario.SelectedItem)
    
        Select Case ObjType
    
            Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otItemsMagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otNudillos, eObjType.otMonturas
                Call frmMain.EquiparObjeto(Inventario.SelectedItem)
                
            Case eObjType.otWeapon
                
                If esArco(Inventario.OBJIndex(Inventario.SelectedItem)) And Inventario.Equipped(Inventario.SelectedItem) Then
                    Call WriteUseItem(Inventario.SelectedItem)
                ElseIf esArrojadiza(Inventario.OBJIndex(Inventario.SelectedItem)) And Inventario.Equipped(Inventario.SelectedItem) Then
                    Call WriteUseItem(Inventario.SelectedItem)
                ElseIf Inventario.OBJIndex(Inventario.SelectedItem) = 15 Then
                    Call WriteUseItem(Inventario.SelectedItem)
                Else
                    Call frmMain.EquiparObjeto(Inventario.SelectedItem)
                End If
                
            Case eObjType.otAnillo
    
                If Inventario.Equipped(Inventario.SelectedItem) Then
                    Call esHerramienta(Inventario.OBJIndex(Inventario.SelectedItem))
                Else
                    Call frmMain.EquiparObjeto(Inventario.SelectedItem)
                End If
                    
            Case Else
            
                Call WriteUseItem(Inventario.SelectedItem)
    
        End Select
    
    End If

End Sub

Public Sub EquiparItem()

    If CurrentUser.Muerto Then Exit Sub

    If Comerciando Then Exit Sub
        
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then Call frmMain.EquiparObjeto(Inventario.SelectedItem)

End Sub

Public Sub AutoUsarU()
 If frmMain.SendTxt.Visible Or AutoUsarActivado = False Then Exit Sub
   'Tecla Usar Item
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mKeyUseObject)) And MainTimer.Check(TimersIndex.UseItemWithU) Then
     If Inventario.ObjType(Inventario.SelectedItem) = eObjType.otpociones Then Call UsarItem
   End If
End Sub

Public Sub AutoUsar()
'--------------------------------------------------
'Author: Shermie80
'Last Modify Date: 19/07/2020
'
'--------------------------------------------------
If frmMain.SendTxt.Visible Then Exit Sub
 
 If AutoUsarActivado Then
   
   'F1
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf1)) Then
    If MacroList(1).mTipe = 4 Then
        Call UsarMacro(1)
    End If
   End If
   
   'F2
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf2)) Then
    If MacroList(2).mTipe = 4 Then
        Call UsarMacro(2)
    End If
   End If
   
   'F3
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf3)) Then
    If MacroList(3).mTipe = 4 Then
        Call UsarMacro(3)
    End If
   End If
   
   'F4
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf4)) Then
    If MacroList(4).mTipe = 4 Then
        Call UsarMacro(4)
    End If
   End If
   
   'F5
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf5)) Then
    If MacroList(5).mTipe = 4 Then
        Call UsarMacro(5)
    End If
   End If
   
   'F6
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf6)) Then
    If MacroList(6).mTipe = 4 Then
        Call UsarMacro(6)
    End If
   End If
   
   'F7
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf7)) Then
    If MacroList(7).mTipe = 4 Then
        Call UsarMacro(7)
    End If
   End If
   
   'F8
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf8)) Then
    If MacroList(8).mTipe = 4 Then
        Call UsarMacro(8)
    End If
   End If
   
   'F9
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf9)) Then
    If MacroList(9).mTipe = 4 Then
        Call UsarMacro(9)
    End If
   End If
   
   'F10
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf10)) Then
    If MacroList(10).mTipe = 4 Then
        Call UsarMacro(10)
    End If
   End If
   
   'F11
   If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mkeyf11)) Then
    If MacroList(11).mTipe = 4 Then
        Call UsarMacro(11)
    End If
    
 End If
   
End If
 
End Sub
