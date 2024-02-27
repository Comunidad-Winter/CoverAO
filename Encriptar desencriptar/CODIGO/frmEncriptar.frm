VERSION 5.00
Begin VB.Form frmEncriptar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Encriptar archivos"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Desencriptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   ".ind"
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encriptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Extension"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "En la carpeta encriptar el archivo debe llamarse ""table"""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmEncriptar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private tmpStr As String
Private MapTable(1 To 18, 1 To 23) As Integer
Private f As Integer
Private Archivo As String
Private Extension As String

Private Sub Command1_Click()


'Encritar

If Len(Extension) <= 0 Then
    MsgBox "Extension invalida"
    Exit Sub
End If

f = FreeFile
Archivo = "Balance"
Extension = ".dat"

If FileExist(App.Path & "\file\" & Archivo & Extension, vbNormal) Then

    Open App.Path & "\file\" & Archivo & Extension For Binary As #f
 
        Get #f, , tmpStr

    Close #f

Else

    MsgBox "El archivo no existe"
    
End If

End Sub

Private Sub Command2_Click()

'Encriptar

Dim i As Long
Dim ii As Long
Dim NumMapas As Integer
Dim Contador As Integer
Dim Mapa() As Integer

Dim AuxMapas() As String
Dim strMapa As String

'Mundo1
'strMapa = "77-743-739-740-741-742-86-162-244-252-284-285-286-287-288-289-291-292-323-322-321-320-319-318-317-316-315-314-137-136-135-134-133-132-131-130-324-325-326-327-328-329-330-331-332-333-61-334-335-336-337-338-339-340-373-372-371-370-369-368-367-366-365-364-60-363-362-361-360-359-358-357-374-375-376-377-378-379-380-381-382-66-59-159-160-161-749-383-384-385-402-401-400-399-398-397-396-395-394-65-58-158-157-156-151-150-149-148-403-227-228-229-404-405-406-407-408-67-57-409-410-411-412-413-414-415-442-224-225-226-441-440-439-438-437-68-56-436-435-434-433-432-431-430-"
'strMapa = strMapa & "443-221-222-223-444-445-446-447-448-69-55-450-451-452-453-454-455-456-482-219-218-220-481-480-479-478-76-70-54-477-476-475-474-473-472-471-483-484-217-485-486-487-488-489-490-71-53-491-492-493-494-495-496-497-521-520-216-519-518-517-516-515-514-72-7-513-512-511-510-509-508-507-522-523-215-524-525-526-527-528-529-73-6-530-531-532-533-534-535-536-562-561-214-560-559-558-557-556-75-74-5-555-554-553-552-21-551-550-563-564-213-212-236-235-234-10-9-8-1-11-12-13"
'strMapa = strMapa & "-15-16-17-103-566-567-568-569-570-571-243-242-38-39-2-14-18-19-98-20-101-102-581-582-583-584-585-586-587-241-46-36-3-25-26-27-97-99-100-588-599-600-601-602-603-604-605-240-80-35-4-22-23-24-96-606-607-608-619-620-621-622-623-624-625-626-78-34-32-29-28-94-95-627-628-629-640-641-642-643-644-645-646-647-79-87-31-30-92-93-648-649-650-651-662-663-664-665-666-667-668-669-670-88-89-90-91-671-672-673-674-675-681-682-683-684-685-686-687-688-689-690-691-692-693-694-695-696-697-698-709-710-711-712-713-714-715-716-717-718-719-720-721-722-723-724-725-726"

'Mundo3
'strMapa = "00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-865-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-830-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-831-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-832-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-833-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00"
'strMapa = strMapa & "-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-118-119-206-237-238-245-246-247-248-249-251-258-00-00-00-00-00-00-290-745-748-750-751-752-838-845-847-848-851-864-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00"

'Mundo2
strMapa = "37-208-00-33-40-00-00-747-00-00-00-00-00-00-00-00-00-49-00-00-00-00-41-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-257-00-42-00-00-00-00-00-00-00-00-00-00-00-00-00-207-00-00-00-43-00-00-00-00-48-00-00-00-00-00-115-00-00-00-00-00-45-44-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-838-00-00-00-00-00-00"
strMapa = strMapa & "-00-50-51-52-00-00-00-00-00-00-00-00-00-00-00-00-210-211-00-00-00-00-00-00-00-00-00-116-00-00-00-00-00-00-00-00-00-753-754-00-758-759-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-239-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-205-00-00-00-00-00-00-00-00-00-00-00-00-254-255-00-00-00-00-00-00-117-00-00-00-233-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-844-00-00-146-00-00-00-00-00-00-231-230-00-00-00-00-00-00-00-00-00-142-00-00-457-00-00-00-232-00-00-00-00-00-00-00-00-00-00-141-143-00-00-00-00-00-250-00-00-00-00-00-00-00-756-757-00-140-144-145-849-850-00-00-"
strMapa = strMapa & "00-00-00-00-00-00-00-00-755-760-00-00-00-00-00-00-00-209-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-204-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-840-841-842-843-852-853-854-855-00-00-00-00-00-00-00-00-00-00-856-857-858-859-860-861-862-863-00-00-00-00-00"

Archivo = "table.ind"

Contador = 0

AuxMapas() = Split(strMapa, "-")

NumMapas = UBound(AuxMapas) + 1
 
ReDim Mapa(1 To NumMapas)

For i = 1 To NumMapas
    Mapa(i) = AuxMapas(i - 1)
     
Next i
 
For i = 1 To 18
    For ii = 1 To 23
        Contador = Contador + 1
        MapTable(i, ii) = Mapa(Contador)
 
    Next ii
Next i

If FileExist(App.Path & "\file\" & Archivo, vbNormal) Then
 
    
     Dim s_ArchivoOrigen As String
   '
   ' Open App.Path & "\file\table.ind" For Binary As #1
   '    s_ArchivoOrigen = Space(LOF(1))
   '    Get #1, , s_ArchivoOrigen
   ' Close #1


    Open App.Path & "\file\grabar.ind" For Binary As #1
       ' Print #1,  MapTable(i, ii)
        
 
For i = 1 To 18
    For ii = 1 To 23
        Contador = Contador + 1
      '  MapTable(i, ii) = Mapa(i * ii)
      
     Put #1, , MapTable(i, ii)

    Next ii
Next i
    Close #1

Else

    MsgBox "El archivo no existe"
    
End If

End Sub
Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean

    FileExist = LenB(Dir$(File, FileType)) <> 0

End Function
Private Sub Text1_Change()
Extension = Text1.Text
End Sub

