Attribute VB_Name = "Resolution"
'*****************************
'Resolucion*******************
'*****************************
'Constantes


Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_TEST = &H4

'Estructura
Public Type DevMode
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public Declare Function EnumDisplaySettings Lib "user32" _
    Alias "EnumDisplaySettingsA" ( _
    ByVal lpszDeviceName As Long, _
    ByVal iModeNum As Long, _
    lpDevMode As Any) As Boolean

Public Declare Function ChangeDisplaySettings Lib "user32" _
    Alias "ChangeDisplaySettingsA" ( _
    lpDevMode As Any, _
    ByVal dwFlags As Long) As Long

Public Declare Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal nIndex As Long) As Long

Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" ( _
    ByVal lpDriverName As String, _
    ByVal lpDeviceName As String, _
    ByVal lpOutput As String, _
    ByVal lpInitData As Any) As Long

Public OldX As Long
Public OldY As Long
Public OldBit As Long

Public ChangeResolution As Boolean


Public Function Engine_Set_Resolution()
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 30/10/2010
'
'**************************************************************

    Dim DevMode As DevMode
    
    
  If RunWindowed = 0 Then

    EnumDisplaySettings 0&, 0&, DevMode
    
 
    If (curDevMode.dmBitsPerPel <> 16) Or (curDevMode.dmPelsHeight <> 600) Or (curDevMode.dmPelsWidth <> 800) Then
    
    OldX = Screen.Width / Screen.TwipsPerPixelX
    OldY = Screen.Height / Screen.TwipsPerPixelY
    
    OldBit = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
    OldBit = GetDeviceCaps(OldBit, 12)
    
    With DevMode
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        .dmPelsWidth = 800 'Ancho de pantalla
        .dmPelsHeight = 600 'Alto de pantalla
        .dmBitsPerPel = Pixels 'Cantidad de bits por pixeles
    End With
    
    ChangeDisplaySettings DevMode, CDS_TEST
    
    ChangeResolution = True
    
    End If
End If
    
End Function


Public Function Engine_Reset_Resolution()
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 30/10/2010
'
'**************************************************************
    If ChangeResolution = False Then Exit Function
    Dim DevMode As DevMode
    
    
If RunWindowed = 0 Then
    If (curDevMode.dmBitsPerPel <> 16) Or (curDevMode.dmPelsHeight <> 600) Or (curDevMode.dmPelsWidth <> 800) Then
    
    EnumDisplaySettings 0&, 0&, DevMode
    
    With DevMode
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        .dmPelsWidth = OldX 'Ancho de pantalla
        .dmPelsHeight = OldY 'Alto de pantalla
        .dmBitsPerPel = OldBit 'cantidad de bits por pixeles
    End With
    
    ChangeDisplaySettings DevMode, CDS_TEST
    End If
End If


End Function

