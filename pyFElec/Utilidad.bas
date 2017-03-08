Attribute VB_Name = "Utilidad"
 'API DECLARATIONS
Declare Function GetPrivateProfileString Lib "kernel32" Alias _
                 "GetPrivateProfileStringA" (ByVal lpApplicationName _
                 As String, ByVal lpKeyName As Any, ByVal lpDefault _
                 As String, ByVal lpReturnedString As String, ByVal _
                 nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias _
                 "WritePrivateProfileStringA" (ByVal lpApplicationName _
                 As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
                 ByVal lpFileName As String) As Long
Public Function sGetINI(sINIFile As String, sSection As String, sKey _
                As String, sDefault As String) As String
    Dim sTemp As String * 256
    Dim nLength As Integer
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, _
              255, sINIFile)
    sGetINI = Left$(sTemp, nLength)
End Function
Public Sub writeINI(sINIFile As String, sSection As String, sKey _
           As String, sValue As String)
    Dim n As Integer
    Dim sTemp As String
    sTemp = sValue
    'Replace any CR/LF characters with spaces
    For n = 1 To Len(sValue)
        If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf _
        Then Mid$(sValue, n) = " "
    Next n
    n = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
End Sub

Sub WriteFile(text As String, filename As String)
    Dim iFileNo As Integer
    iFileNo = FreeFile
    'open the file for writing
    Path = CurDir() + "\"
    Open Path + filename For Output As #iFileNo
    'please note, if this file already exists it will be overwritten!
    
    'write some example text to the file
     Print #iFileNo, text
    'close the file (if you dont do this, you wont be able to open it again!)
    Close #iFileNo
End Sub

Public Sub BuscarRuta(CommonDialog As CommonDialog, volcarEn As textbox, tipoArchivo As String)
' CancelError is True.
   On Error GoTo ErrHandler
   ' Set filters.
   CommonDialog.Filter = tipoArchivo + "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Batch Files (*.bat)|*.bat"
   ' Specify default filter.
   CommonDialog.FilterIndex = 1
   ' Display the Open dialog box.
   CommonDialog.ShowOpen
   ' Call the open file procedure.
   volcarEn.text = CommonDialog.filename
   Exit Sub

ErrHandler:
' User pressed Cancel button.
   Exit Sub
End Sub

Public Function RowIsEmpty(isEmpty As Boolean, grilla As ubGrid, fila As Integer) As Boolean
    Dim i As Integer
    isEmpty = True
    grilla.row = fila
    For i = 0 To grilla.Cols
        If grilla.text <> "" Then
            isEmpty = False
            Exit For
        End If
    Next
    RowIsEmpty = isEmpty
End Function

Public Function GetNumberFromString(texto As String) As String
    If texto = "" Then
        GetNumberFromString = 0
    Else
        GetNumberFromString = texto
    End If
End Function

Public Function CentrarForm(frm As Form)
    frm.Move MDIMain.Width / 2 - frm.Width / 2, (MDIMain.Height - 700) / 2 - frm.Height / 2
End Function

Public Function IfNotExistsCreateFile(volcarEn As String, ini As String, rama As String, clave As String, valorDefault As String) As String
    volcarEn = sGetINI(ini, rama, clave, "?")
    If volcarEn = "?" Then
        writeINI ini, rama, clave, valorDefault
        IfNotExistsCreateFile = valorDefault
    Else
        IfNotExistsCreateFile = volcarEn
    End If
End Function
Public Function RecuperarContenidoTextbox(ByRef textbox As textbox, recupero As String)
    If recupero <> "?" Then
          textbox.text = recupero
    End If
End Function
