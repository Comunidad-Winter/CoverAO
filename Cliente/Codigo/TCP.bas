Attribute VB_Name = "Mod_TCP"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public LlegaronSkills As Boolean
Public LlegaronStats  As Boolean
Public LlegaronAtrib  As Boolean

Sub Login()
    
    Select Case EstadoLogin
    
    
    Case E_MODO.Normal
        Call WriteLoginAccount
        
    Case E_MODO.CrearNuevoPj
        Call WriteLoginNewChar
        
    Case E_MODO.CrearNuevaCuenta
        Call WriteLoginNewAccount
    
    Case E_MODO.RecuperarCuenta, E_MODO.CambiarContrase�a, E_MODO.BorrarPersonaje
        Call WriteProcesosLogin
        
    Case E_MODO.ConectarPersonaje
        Call WriteLoginExistingChar

    End Select
    DoEvents

    Call FlushBuffer
    
End Sub

