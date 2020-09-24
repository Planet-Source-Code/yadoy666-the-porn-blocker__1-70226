Attribute VB_Name = "Fungsi"
Public Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
 
Public RByte()     As Variant
Public Const WM_CLOSE = &H10

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type


Enum SFolder
    CSIDL_PROGRAMS = &H2
End Enum

Public IP As String, Situs As String
Public x As String, Judul As String


Public Sub LoadFileHost(list As ListBox, Namafile As String)
Dim linestr As String, tmp() As String
On Error Resume Next
    Open Namafile For Input As #1
        While Not EOF(1)
          Line Input #1, linestr
        tmp = Split(linestr, "       ")
        IP = tmp(0)
        Situs = tmp(1)
        DoEvents
            list.AddItem Situs
        Wend
    Close #1
End Sub
Public Sub Load_Caption(list As ListBox, Namafile As String)
Dim linestr As String, tmp() As String
On Error Resume Next
Open Namafile For Input As #1
    While Not EOF(1)
        Line Input #1, linestr
        Judul = linestr
    DoEvents
        list.AddItem Judul
    Wend
Close #1
End Sub


Public Sub SaveFileHost(list As ListBox, place As String)
On Error Resume Next
Dim simpan As Long
    Open place For Output As #1
        For simpan = 0 To list.ListCount - 1
            Print #1, "127.0.0.1       " & list.list(simpan)
        Next
    Close #1
End Sub
Public Sub SaveCaption(list As ListBox, place As String)
On Error Resume Next
Dim simpan As Long
    Open place For Output As #1
        For simpan = 0 To list.ListCount - 1
            Print #1, list.list(simpan)
        Next
Close #1
End Sub


Public Sub hapus(list As ListBox, place As String)
On Error Resume Next
Dim hapus As Long
Open place For Output As #1
  For hapus = 0 To list.ListCount - 1
    Print #1, "127.0.0.1       " & list.list(hapus)
  Next
Close #1
End Sub

Public Sub HapusCaption(list As ListBox, place As String)
On Error Resume Next
Dim hapus As Long
Open place For Output As #1
    For hapus = 0 To list.ListCount - 1
        Print #1, list.list(hapus)
    Next
Close #1
End Sub


Public Sub backup()
FileCopy GetSystemPath & "\Drivers\etc\Hosts", App.Path & "\back.txt"
Open GetSystemPath & "\Drivers\etc\Hosts" For Output As #1
    Print #1, "127.0.0.1          localhost"
Close #1
End Sub
Public Sub mulai()
On Error Resume Next
FileCopy App.Path & "\back.txt", GetSystemPath & "\Drivers\etc\Hosts"
FileCopy App.Path & "back.txt", GetSystemPath & "\Drivers\etc\Hosts"

End Sub


Public Function GetSystemPath() As String

On Error Resume Next
Dim Buffer As String * 255
Dim x As Long
    x = GetSystemDirectory(Buffer, 255)
    GetSystemPath = Left(Buffer, x) & "\"
    
End Function

Public Function Hajar(target As String)
Dim h As Long
Dim t As String * 255
h = GetForegroundWindow
GetWindowText h, t, 255
If InStr(UCase(t), UCase(target)) > 0 Then
    SendMessage h, WM_CLOSE, 0, 0
MsgBox "Maaf perintah yang coba anda jalankan telah dinonaktifkan oleh administrator komputer ini. Silahkan menghubungi administrator untuk mengaktifkannya kembali", vbInformation + vbOKOnly, "Pembatasan"
End If
End Function


Public Sub Tonjok(target As String)
Dim h As Long
Dim t As String * 255
h = GetForegroundWindow
GetWindowText h, t, 255

If InStr(UCase(t), UCase(target)) > 0 Then
    SendMessage h, WM_CLOSE, 0, 0
End If
End Sub

Public Sub kill_IE(target As String)
Dim h As Long
Dim t As String * 255
h = GetForegroundWindow
GetWindowText h, t, 255

If InStr(UCase(t), UCase(target)) > 0 Then
Shell App.Path & "\kill.bat", vbHide
End If
End Sub

Public Function GetSpecialfolder(JenisFolder As SFolder) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    'get special folder
    r = SHGetSpecialFolderLocation(100, JenisFolder, IDL)
    If r = NOERROR Then
        'create buffer
        Path$ = Space$(512)
        'Get path from IDList(IDL)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        'Remove chr$(0)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Public Function Getprogramfile() As String
    
    On Error Resume Next
    Dim r As Long
    Dim IDL As ITEMIDLIST
    Dim I As Integer
    'dapatkan special folder
    r = SHGetSpecialFolderLocation(100, CSIDL_PROGRAMS, IDL)
    If r = NOERROR Then
        'buat buffer
        Path$ = Space$(512)
        'dapatkan path dari IDList(IDL)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        'hapus karakter chr$(0) yang tidak dibutuhkan
        Getprogramfile = Left$(Path, InStr(Path, Chr$(0)) - 1)

        For I = Len(Getprogramfile) To 1 Step -1
            If Mid$(Getprogramfile, I, 1) = "\" Then
                Getprogramfile = Left(Getprogramfile, I - 1) & "\Porn_Blocker"
                Exit Function
            End If
            DoEvents
        Next I
        
        Exit Function
    End If
    Getprogramfile = ""
    
End Function

Public Function crypt(strInput As String, _
                       ByVal bEncrypt As Boolean) As String

  Dim I      As Long
  Dim NewAsc As Long
  Dim keypos As Long

  GetKey 1
  For I = 1 To Len(strInput)
    If bEncrypt Then
      NewAsc = Asc(Mid$(strInput, I, 1)) + RByte(keypos)
     Else
      NewAsc = Asc(Mid$(strInput, I, 1)) - RByte(keypos)
    End If
    Do While NewAsc < 0
      NewAsc = NewAsc + 255
    Loop
    Do While NewAsc > 255
      NewAsc = NewAsc - 255
    Loop
    Mid$(strInput, I, 1) = Chr$(NewAsc)
    keypos = keypos + 1
    If keypos > UBound(RByte) Then
      keypos = 0
    End If
  Next I
  crypt = strInput

End Function
Public Sub GetKey(StrA As String)

  Dim I As Long
 If Len(StrA) Then
  ReDim RByte(Len(StrA) - 1) As Variant
  For I = 1 To Len(StrA)
    RByte(I - 1) = Asc(Mid$(StrA, I, 1))
  Next I
End If
End Sub


