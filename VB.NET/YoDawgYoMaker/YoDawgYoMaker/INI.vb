Imports System.Object
Imports System.Runtime.InteropServices.Marshal

Module ComINI

    Private Declare Function GetPrivateProfileSection Lib "kernel32.dll" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedBuffer As IntPtr, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As IntPtr, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedBuffer As IntPtr, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function WritePrivateProfileSection Lib "kernel32.dll" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long




    'Retourne un booléen indiquant l'existance ou non d'une section
    Public Function ExisteSection(ByVal File As String, ByVal Section As String) As Boolean
        Dim PtrCh As IntPtr
        Dim Lng As Integer

        PtrCh = StringToHGlobalAnsi(New String(vbNullChar, 1024))
        Lng = GetPrivateProfileSection(Section, PtrCh, 1024, File)

        Return Lng
    End Function

    'Retourne une collection contenant l'ensemble des sections du fichier "File"
    Public Function GetAllSections(ByVal File As String) As Collection
        Dim PtrCh As IntPtr
        Dim Sections As Collection
        Dim I, Lng As Integer
        Dim Chaine, SChaine As String

        PtrCh = StringToHGlobalAnsi(New String(vbNullChar, 1024))
        Lng = GetPrivateProfileSectionNames(PtrCh, 1024, File)
        Chaine = PtrToStringAnsi(PtrCh, Lng)
        FreeHGlobal(PtrCh)

        Sections = New Collection
        SChaine = ""
        For I = 0 To Lng - 1
            If Chaine.Chars(I) = vbNullChar Then
                Sections.Add(SChaine)
                SChaine = ""
            Else
                SChaine = SChaine & Chaine.Chars(I)
            End If
        Next
        GetAllSections = Sections
        Sections = Nothing
    End Function

    'Retourne une collection contenant l'ensemble des clés de la section "Section" du fichier "File"
    Public Function GetSectionCles(ByVal File As String, ByVal Section As String) As Collection
        Dim PtrCh As IntPtr
        Dim Cles As Collection
        Dim I, Lng As Integer
        Dim Chaine, SChaine As String

        PtrCh = StringToHGlobalAnsi(New String(vbNullChar, 1024))
        Lng = GetPrivateProfileSection(Section, PtrCh, 1024, File)
        Chaine = PtrToStringAnsi(PtrCh, Lng)
        FreeHGlobal(PtrCh)

        Cles = New Collection
        SChaine = ""
        For I = 0 To Lng - 1
            If Chaine.Chars(I) = vbNullChar Then
                Cles.Add(SChaine)
                SChaine = ""
            Else
                SChaine = SChaine & Chaine.Chars(I)
            End If
        Next
        GetSectionCles = Cles
        Cles = Nothing
    End Function

    'Retourne la valeur de la clé "Cle" de la section "Section" du fichier "File"
    Public Function GetCle(ByVal File As String, ByVal Section As String, ByVal Cle As String) As String
        Dim PtrCh As IntPtr
        Dim Lng As Integer
        Dim Chaine As String

        PtrCh = StringToHGlobalAnsi(New String(vbNullChar, 1024))
        Lng = GetPrivateProfileString(Section, Cle, "", PtrCh, 255, File)
        Chaine = PtrToStringAnsi(PtrCh, Lng)
        FreeHGlobal(PtrCh)

        GetCle = Chaine
    End Function


    'Insère une section dans le fichier "File"
    Public Function SetSection(ByVal File As String, ByVal Section As String, ByVal Valeur As String) As Boolean
        SetSection = WritePrivateProfileSection(Section, Valeur, File)
    End Function

    'Insère la clé "Cle" dans la section "Section" du fichier "File"
    Public Function SetCle(ByVal File As String, ByVal Section As String, ByVal Cle As String, ByVal Valeur As String) As Boolean
        SetCle = WritePrivateProfileString(Section, Cle, Valeur, File)
    End Function


    'Efface toute les clés de la section "Section"
    Public Function DelSection(ByVal File As String, ByVal Section As String) As Boolean
        DelSection = WritePrivateProfileSection(Section, "", File)
    End Function

    'Efface la valeur de la clé "Cle" de la section "Section"
    Public Function DelCle(ByVal File As String, ByVal Section As String, ByVal Cle As String) As Boolean
        DelCle = WritePrivateProfileString(Section, Cle, "", File)
    End Function

End Module
