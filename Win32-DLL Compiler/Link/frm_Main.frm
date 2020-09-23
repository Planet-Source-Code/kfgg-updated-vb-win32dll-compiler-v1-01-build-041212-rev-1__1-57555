VERSION 5.00
Begin VB.Form frm_Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB-Win32DLL Compiler"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5025
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame fra_Container 
      Caption         =   "Operation"
      Height          =   1275
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   3240
      Width           =   4830
      Begin VB.CommandButton cmd_About 
         Caption         =   "&About"
         Height          =   420
         Left            =   3240
         TabIndex        =   18
         Top             =   720
         Width           =   1500
      End
      Begin VB.CommandButton cmd_Parameter 
         Caption         =   "Parameter for no&rmal compile"
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   720
         Width           =   1500
      End
      Begin VB.CommandButton cmd_Parameter 
         Caption         =   "Parameter for Win32 D&LL"
         Height          =   420
         Index           =   1
         Left            =   1665
         TabIndex        =   16
         Top             =   720
         Width           =   1500
      End
      Begin VB.CommandButton cmd_Compile 
         Caption         =   "Compile as &parameter"
         Height          =   420
         Index           =   2
         Left            =   3240
         TabIndex        =   12
         Top             =   225
         Width           =   1500
      End
      Begin VB.CommandButton cmd_Compile 
         Caption         =   "Compile as &normal"
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   1500
      End
      Begin VB.CommandButton cmd_Compile 
         Caption         =   "Compile Win32 &DLL"
         Height          =   420
         Index           =   1
         Left            =   1665
         TabIndex        =   2
         Top             =   225
         Width           =   1500
      End
   End
   Begin VB.Frame fra_Container 
      Caption         =   "Compiler settings"
      Height          =   3030
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4830
      Begin VB.TextBox txt_FileName 
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   14
         Top             =   2565
         Width           =   4065
      End
      Begin VB.CommandButton cmd_FileOpt 
         Height          =   285
         Index           =   1
         Left            =   4320
         Picture         =   "frm_Main.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1980
         Width           =   375
      End
      Begin VB.CommandButton cmd_FileOpt 
         Height          =   285
         Index           =   2
         Left            =   4320
         Picture         =   "frm_Main.frx":0894
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2565
         Width           =   375
      End
      Begin VB.CommandButton cmd_FileOpt 
         Height          =   285
         Index           =   0
         Left            =   4320
         Picture         =   "frm_Main.frx":0E1E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1395
         Width           =   375
      End
      Begin VB.TextBox txt_FileName 
         Height          =   285
         Index           =   1
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1980
         Width           =   4065
      End
      Begin VB.TextBox txt_Command 
         Height          =   600
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   5
         Top             =   495
         Width           =   4560
      End
      Begin VB.TextBox txt_FileName 
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Text            =   "LinkMS.exe"
         Top             =   1395
         Width           =   4065
      End
      Begin VB.Label lbl_Tips 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File to exporting declares"
         Height          =   180
         Index           =   4
         Left            =   135
         TabIndex        =   15
         Top             =   2340
         Width           =   2340
      End
      Begin VB.Label lbl_Tips 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Files that contain functions or subs to export"
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   1755
         Width           =   4140
      End
      Begin VB.Label lbl_Tips 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File of the compiler (The renamed Link.exe)"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   1170
         Width           =   3870
      End
      Begin VB.Label lbl_Tips 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parameter"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   270
         Width           =   810
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***VB-Win32DLL Compiler, made by KFGG, China.P.R
'***V1.01 build 041212 (Rev. 1)
'***12/12/2004
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Long) As Integer
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lstructSize As Long
    hwnd As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10
Private Const VK_SHIFT = &H10

Private Const DLLNameMark = "/OUT:"
Private Const Version = "V1.01 build 041212"
Private Const AppName = "VB-Win32DLL Compiler"
Private Const CopyRight = "(C) KFGG, China.P.R, 2004"
Dim ExportParameter As String, DLLName As String, Parameter As String
'Const Command$ = """G:\VB Projects\qqbot-m\Class1.OBJ"" ""G:\VB Projects\qqbot-m\PRO1.OBJ"" ""d:\vb6\vb98\VBAEXE6.LIB"" /ENTRY:__vbaS /OUT:""G:\VB Projects\qqbot-m\¹¤³Ì1.dll"" /BASE:0x400000 /SUBSYSTEM:WINDOWS,4.0 /VERSION:1.0   /INCREMENTAL:NO /OPT:REF /MERGE:.rdata=.text /IGNORE:4078 """

Function DialogFile(hwnd As Long, wMode As Integer, szDialogTitle As String, szFilename As String, szFilter As String, szDefDir As String, szDefExt As String, Optional MultiSelect As Boolean = False) As String
'open and save common dialog

'i have forgotten the author of this code(lost the original full code),
'but i am sure that i have got it from PSC,anyone who had made this code
'please mail me so that i can add your name to the reference

    Dim x As Long, OFN As OPENFILENAME, szFile As String, szFileTitle As String
    
    OFN.lstructSize = Len(OFN)
    OFN.hwnd = hwnd
    OFN.lpstrTitle = szDialogTitle
    OFN.lpstrFile = szFilename & String$(250 - Len(szFilename), 0)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = String$(255, 0)
    OFN.nMaxFileTitle = 255
    OFN.lpstrFilter = szFilter
    OFN.nFilterIndex = 1
    OFN.lpstrInitialDir = szDefDir
    OFN.lpstrDefExt = szDefExt
    If MultiSelect Then OFN.Flags = OFN.Flags Or OFN_ALLOWMULTISELECT
    If wMode = 1 Then
        OFN.Flags = OFN.Flags Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        x = GetOpenFileName(OFN)
    Else
        OFN.Flags = OFN.Flags Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
        x = GetSaveFileName(OFN)
    End If
    If x <> 0 Then
        '// If Instr(OFN.lpstrFileTitle, Chr$(0)) > 0 Then
        '//     szFileTitle = Left$(OFN.lpstrFileTitle, Instr(OFN.lpstrFileTitle, Chr$(0)) - 1)
        '// End If
        If InStr(OFN.lpstrFile, Chr$(0)) > 0 Then
            szFile = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr$(0)) - 1)
        End If
        '// OFN.nFileOffset is the number of characters from the beginning of the
        '// full path to the start of the file name
        '// OFN.nFileExtension is the number of characters from the beginning of the
        '// full path to the file's extention, including the (.)
        '// MsgBox "File Name is " & szFileTitle & Chr$(13) & Chr$(10) & "Full path and file is " & szFile, , "Open"
        
        '// DialogFile = szFile & "|" & szFileTitle
        DialogFile = szFile
    Else
        DialogFile = ""
    End If
End Function
Function Press(vKeyCode As Long) As Boolean
'capture the key you pressed
  Press = (GetAsyncKeyState(vKeyCode) < 0)
End Function
Sub AlwaysOnTop(hwnd As Long, Always_On_Top As Boolean)
  If Always_On_Top Then
    SetWindowPos hwnd, -1, 0, 0, 0, 0, 3
  Else
    SetWindowPos hwnd, -2, 0, 0, 0, 0, 3
  End If
End Sub

Function CheckFunc(ByVal strToCheck As String) As Long
'0-not func or sub,1-sub,2-function,3,4-public...
  CheckFunc = 0
  If (Left$(strToCheck, 3) = "Sub") Then
    CheckFunc = 1
    Exit Function
  End If
  If (Left$(strToCheck, 8) = "Function") Then
    CheckFunc = 2
    Exit Function
  End If
  If (Left$(strToCheck, 10) = "Public Sub") Then
    CheckFunc = 3
    Exit Function
  End If
  If (Left$(strToCheck, 15) = "Public Function") Then
    CheckFunc = 4
    Exit Function
  End If
End Function
Function GetDLLName() As String
'get the filename you want to compile as from command
Dim s1 As String, s2 As String, s3 As String, i As Long
On Error Resume Next
  s1 = Mid$(Parameter, InStr(1, Parameter, DLLNameMark, vbTextCompare) + Len(DLLNameMark) + 1, Len(Parameter))
  s1 = Left$(s1, InStr(1, s1, """", vbTextCompare) - 1)
  i = Len(s1)
  If i = 0 Then Exit Function
  Do
    s2 = s2 + s3
    s3 = Mid$(s1, i, 1)
    i = i - 1
  Loop Until s3 = "\"
  s2 = StrReverse(s2)
  GetDLLName = s2
End Function
Function MakeParameter() As String
'replace "exe project" parameter by "dll project"
Dim strToMake As String
  strToMake = Replace(Command$, "0x400000", "0x11000000", 1, -1, vbTextCompare)
  'If InStr(1, strToMake, "/DLL  /INCREMENTAL", vbTextCompare) = 0 Then strToMake = Replace(strToMake, "  /INCREMENTAL", "/DLL  /INCREMENTAL", 1, -1, vbTextCompare)
  MakeParameter = strToMake
End Function
Function MakeDeclare(ByVal strToMake As String, ByVal strType As Long) As String
'make declare for exported functions and subs
Dim strDeclare As String, strFuncName As String, strFuncCmd As String, i As Long
  Select Case strType
  
  Case 1
     strDeclare = "Declare Sub "
     strFuncName = Mid$(strToMake, 5, Len(strToMake))
  Case 2
     strDeclare = "Declare Function "
     strFuncName = Mid$(strToMake, 10, Len(strToMake))
  Case 3
     strDeclare = "Declare Sub "
     strFuncName = Mid$(strToMake, 12, Len(strToMake))
  Case 4
     strDeclare = "Declare Function "
     strFuncName = Mid$(strToMake, 17, Len(strToMake))
  End Select
  
  i = InStr(1, strFuncName, "(")
  strFuncCmd = Mid$(strFuncName, i, Len(strFuncName))
  strFuncName = Left$(strFuncName, i - 1)
  strDeclare = strDeclare + strFuncName + " Lib " + """" + DLLName + """" + " Alias " + """" + strFuncName + """" + " " + strFuncCmd
  ExportParameter = ExportParameter + "/EXPORT:" + strFuncName + " "
  'make parameter required exporting functions and subs
  strDeclare = LeftB$(strDeclare, LenB(strDeclare) - 2)
  MakeDeclare = strDeclare
End Function

Function CompileDLL() As String
'do all needed for a exportable dll or exe
Dim FileNo1 As Long, FileNo2 As Long, FileNo3 As Long, ChrRead As Byte, strToWrite As String, s1 As String, s2 As String, i As Long
On Error Resume Next
  ExportParameter = ""
  Kill Me.txt_FileName(2).Text
  FileNo2 = FreeFile
  Open Me.txt_FileName(2).Text For Binary As FileNo2
  FileNo3 = FreeFile
  Open Me.txt_FileName(2).Text + "~" For Binary As FileNo3
    Put #FileNo2, , DLLName + " declares made by " + AppName + " at " + Date$ + " " + Time$ + vbCrLf
    'Put #FileNo2, , CopyRight + vbCrLf
    Put #FileNo2, , "Public Declares:" + vbCrLf
    Put #FileNo3, , "Private Declares:" + vbCrLf
    s1 = Me.txt_FileName(1).Text
    i = InStr(1, s1, " ", vbTextCompare)
    s1 = Mid$(s1, i + 1, Len(s1))
  Do 'get all modules' file name
    i = InStr(1, s1, " ", vbTextCompare)
    If i = 0 Then
      s2 = s1
    Else
      s2 = Left$(s1, i - 1)
      s1 = Mid$(s1, i + 1, Len(s1))
    End If
  
    FileNo1 = FreeFile
    Open s2 For Binary As FileNo1
    'open a module or text file that contains declares
      Seek FileNo1, 1
      Do
        Get FileNo1, , ChrRead
        If ChrRead = 10 Then 'if end of current line
          If Mid$(strToWrite, Len(strToWrite) - 1, 1) <> "_" Then
          'if it is a multi line function or sub
            If CheckFunc(strToWrite) <> 0 Then
            'if the line is the start of a function or sub
              strToWrite = MakeDeclare(strToWrite, CheckFunc(strToWrite))
              Put #FileNo2, , "Public " + strToWrite + vbCrLf 'public declare
              Put #FileNo3, , "Private " + strToWrite + vbCrLf 'private declare
            End If
            strToWrite = ""
          Else
            strToWrite = Left$(strToWrite, (Len(strToWrite) - 2))
          End If
        Else
          strToWrite = strToWrite + Chr(ChrRead)
        End If
      Loop Until EOF(FileNo1)
    Close FileNo1
  Loop Until i = 0
  Close FileNo2
  Close FileNo3
  
  FileNo2 = FreeFile
  Open Me.txt_FileName(2).Text For Binary As FileNo2
  FileNo3 = FreeFile
  Open Me.txt_FileName(2).Text + "~" For Binary As FileNo3
    Seek FileNo2, LOF(FileNo2) + 1
    Seek FileNo3, 1
    Put FileNo2, , vbCrLf
    Do
      Get FileNo3, , ChrRead
      Put FileNo2, , ChrRead
    Loop Until EOF(FileNo3)
  Close FileNo3
  Close FileNo2
  'join public and private declare file together
  Kill Me.txt_FileName(2).Text + "~" 'delete temporary file
  i = InStr(1, Parameter, DLLNameMark, vbTextCompare)
  ExportParameter = Left$(Parameter, i - 1) + ExportParameter
  ExportParameter = ExportParameter + Mid$(Parameter, i, Len(Parameter))
  'make parameter requied to a exportable dll or exe
  CompileDLL = ExportParameter
End Function

Private Sub cmd_About_Click()
  MsgBox AppName + " " + Version + vbCrLf + vbCrLf + "Compile a exportable DLL or EXE" + vbCrLf + vbCrLf + CopyRight, vbInformation, AppName
End Sub

Private Sub cmd_FileOpt_Click(Index As Integer)
Dim i As Long, strFileName As String, s1 As String
  Select Case Index
  Case 0 'select compiler(the renamed link.exe)
    strFileName = DialogFile(Me.hwnd, 1, "Files of compiler", "", "Executable file(*.exe)" & Chr(0) & "*.exe", App.Path, "exe")
    If strFileName <> "" Then Me.txt_FileName(0).Text = strFileName
  Case 1 'select modules including the functions and subs you want to export
    strFileName = DialogFile(Me.hwnd, 1, "Files that contain functions or subs to export", "", "VB Module files(*.bas)" & Chr(0) & "*.bas" & Chr(0) & "Text files(*.txt)" & Chr(0) & "*.txt" & Chr(0) & "All files(*.*)" & Chr(0) & "*.*", Me.txt_FileName(1).Text, "bas", True)
    If strFileName <> "" Then
      Me.txt_FileName(1).Text = strFileName
      i = InStr(1, strFileName, " ", vbTextCompare)
      If i <> 0 Then
        s1 = Left$(strFileName, i - 1)
        strFileName = StrReverse(strFileName)
        i = InStr(1, strFileName, " ", vbTextCompare)
        strFileName = Left$(strFileName, i - 1)
        strFileName = s1 + StrReverse(strFileName)
      End If
      i = Len(strFileName)
      Do
        s1 = Mid$(strFileName, i, 1)
        i = i - 1
      Loop Until s1 = "."
      strFileName = Left$(strFileName, i) + ".txt"
      Me.txt_FileName(2).Text = strFileName
    Else
      Exit Sub
    End If
  Case 2 'select the file you want to export the declares
    strFileName = DialogFile(Me.hwnd, 2, "File of exporting declares", "", "Text file(*.txt)" & Chr(0) & "*.txt" & Chr(0) & "All files(*.*)" & Chr(0) & "*.*", Me.txt_FileName(2).Text, "txt")
    If strFileName <> "" Then Me.txt_FileName(2).Text = strFileName
  End Select
End Sub

Private Sub cmd_Compile_Click(Index As Integer)
  Select Case Index
  Case 0 'normal compile,do just as vb do
    Shell Me.txt_FileName(0).Text + " " + Command$
  Case 1 'compile a exportable dll or exe
    If Me.txt_FileName(1).Text = "" Or Me.txt_FileName(2).Text = "" Then
      MsgBox """Files of modules of your project"" or ""File of exporting declares"" cannot be left empty", vbCritical, AppName
      Exit Sub
    End If
      Shell Me.txt_FileName(0).Text + " " + CompileDLL
  Case 2 'compile as the parameter you specify
    Shell Me.txt_FileName(0).Text + " " + Me.txt_Command.Text
  End Select
  End
End Sub

Private Sub cmd_Parameter_Click(Index As Integer)
  Select Case Index
  Case 0 'show the original parameter
    Me.txt_Command.Text = Command$
  Case 1 'show the parameter for a exportable dll or exe
    If Me.txt_FileName(1).Text = "" Or Me.txt_FileName(2).Text = "" Then
      MsgBox """Files of modules of your project"" or ""File of exporting declares"" cannot be left empty", vbCritical, AppName
      Exit Sub
    End If
    Me.txt_Command.Text = CompileDLL
  End Select
  
End Sub

Private Sub Form_Load()
  If Command$ = "" Then 'check parameter
    MsgBox "Cannot compile with no parameter", vbCritical, AppName
    End
  Else
    If Press(VK_SHIFT) Then 'check if SHIFT pressed
      Parameter = MakeParameter
      Me.txt_Command.Text = Command$
      DLLName = GetDLLName
      AlwaysOnTop Me.hwnd, True
    Else
      cmd_Compile_Click 0
    End If
  End If
End Sub

Private Sub txt_Command_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then Me.txt_Command.Text = Replace(Me.txt_Command.Text, vbCrLf, "", 1, -1, vbTextCompare)
End Sub
