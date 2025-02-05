Attribute VB_Name = "ModEncrypt"

Public Function SEncriptar(ByVal Cadena As String) As String
        
        On Error GoTo SEncriptar_Err
        

        ' GSZ-AO - Encripta una cadena de texto
        Dim i As Long, RandomNum As Integer
    
100     RandomNum = 99 * Rnd

102     If RandomNum < 10 Then RandomNum = 10

104     For i = 1 To Len(Cadena)
106         Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) + RandomNum)
108     Next i

110     SEncriptar = Cadena & Chr$(Asc(Left$(RandomNum, 1)) + 10) & Chr$(Asc(Right$(RandomNum, 1)) + 10)
        'DoEvents (WyroX: WTF?)

        
        Exit Function

SEncriptar_Err:
112     Call RegistrarError(Err.Number, Err.description, "ModEncrypt.SEncriptar", Erl)
114     Resume Next
        
End Function

Public Function SDesencriptar(ByVal Cadena As String) As String
        
        On Error GoTo SDesencriptar_Err
        

        ' GSZ-AO - Desencripta una cadena de texto
        Dim i As Long, NumDesencriptar As String
    
100     NumDesencriptar = Chr$(Asc(Left$((Right(Cadena, 2)), 1)) - 10) & Chr$(Asc(Right$((Right(Cadena, 2)), 1)) - 10)
102     Cadena = (Left$(Cadena, Len(Cadena) - 2))

104     For i = 1 To Len(Cadena)
106         Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) - NumDesencriptar)
108     Next i

110     SDesencriptar = Cadena
        'DoEvents (WyroX: WTF?)

        
        Exit Function

SDesencriptar_Err:
112     Call RegistrarError(Err.Number, Err.description, "ModEncrypt.SDesencriptar", Erl)
114     Resume Next
        
End Function

' GSZAO - Encriptación basica y rapida para Strings
Public Function RndCrypt(ByVal str As String, ByVal Password As String) As String
        
        On Error GoTo RndCrypt_Err
        

        '  Made by Michael Ciurescu
        ' (CVMichael from vbforums.com)
        '  Original thread: http://www.vbforums.com/showthread.php?t=231798
        Dim SK As Long, K As Long

100     Rnd -1
102     Randomize Len(Password)

104     For K = 1 To Len(Password)
106         SK = SK + (((K Mod 256) Xor Asc(mid$(Password, K, 1))) Xor Fix(256 * Rnd))
108     Next K

110     Rnd -1
112     Randomize SK
    
114     For K = 1 To Len(str)
116         Mid$(str, K, 1) = Chr(Fix(256 * Rnd) Xor Asc(mid$(str, K, 1)))
118     Next K
    
120     RndCrypt = str

        
        Exit Function

RndCrypt_Err:
122     Call RegistrarError(Err.Number, Err.description, "ModEncrypt.RndCrypt", Erl)
124     Resume Next
        
End Function

Function RandomString(cb As Integer, Optional ByVal OnlyUpper As Boolean = False) As String
        
        On Error GoTo RandomString_Err
        

100     Randomize Time

        Dim rgch As String

102     rgch = "abcdefghijklmnopqrstuvwxyz"
    
104     If OnlyUpper Then
106         rgch = UCase(rgch)
        Else
108         rgch = rgch & UCase(rgch)

        End If
    
110     rgch = rgch & "0123456789"  ' & "#@!~$()-_"

        Dim i As Long

112     For i = 1 To cb
114         RandomString = RandomString & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Next

        
        Exit Function

RandomString_Err:
116     Call RegistrarError(Err.Number, Err.description, "General.RandomString", Erl)
118     Resume Next
        
End Function


