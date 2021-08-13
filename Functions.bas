'Contents

'    Public Declare Variables
'A!: Public Declare Functions and function for Debug
'B!: General Functions
    'B1!: Maths Functions
    'B2!: Statistics Functions
    'B3!: String Functions
'C!: Array Functions
    'C1!: 1-Dimension Array Functions
    'C2!: 2-Dimension array Functions
        'C2a: Matrix Functions
    'C4!: Jagged Array Functions
    'C3!: 3-Dimension array Functions
'D!: Functions for objects
    'D1! Office Application
'E!: Functions for excel
'F!: Functions for access
'G!: Functions for Word
'H!: Functions for Powerpoint
'I!: Functions for Outlook
'J!: Other Functions
    'J1!: Mouse & Keybroad Control Functions
    
'******************************************************************************************************************************************************

'Public Declare variables
Option Explicit

'Variables
'Variables used in enumwin Function
Public Win() As Variant
Public i As Long
Public ct As Long
Public rect As Boolean

'Types
Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Type POINT
   Xcoord As Long
   Ycoord As Long
End Type

'Declare a UDT to store a GUID for the IPicture OLE Interface
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
 
'Declare a UDT to store the bitmap information
Private Type uPicDesc
    size As Long
    Type As Long
    hPic As Long
    hPal As Long
End Type

'Public Variables

'Constants
'Maths Constants

'System Constants
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000
Private Const GWL_ID As Long = -12
Private Const MIN_BOX As Long = &H20000
Private Const MAX_BOX As Long = &H10000
Private Const SC_RESTORE As Long = &HF120
Private Const SC_CLOSE As Long = &HF060
Private Const SC_MAXIMIZE As Long = &HF030
Private Const SC_MINIMIZE As Long = &HF020
Private Const LEFTDOWN = &H2
Private Const LEFTUP = &H4
Private Const RIGHTDOWN As Long = &H8
Private Const RIGHTUP As Long = &H10
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SNAPSHOT = &H2C
Private Const VK_MENU = &H12

'The API format types we're interested in
Public Const CF_BITMAP = 2
Public Const CF_PALETTE = 9
Public Const CF_ENHMETAFILE = 14
Public Const IMAGE_BITMAP = 0
Public Const LR_COPYRETURNORG = &H4

'Type Time
 Private Type FILETIME
     dwLowDate  As Long
     dwHighDate As Long
 End Type
 
 Private Type SYSTEMTIME
     wYear      As Integer
     wMonth     As Integer
     wDayOfWeek As Integer
     wDay       As Integer
     wHour      As Integer
     wMinute    As Integer
     wSecond    As Integer
     wMillisecs As Integer
 End Type
 
'A!: Public Declare Functions and function for Debug

'API Functions
'VBA7 Declaration
'#if VBA7 Then
'Windows Functions
'Public Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long
'Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal lngHWnd As Long) As Long
'Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Public Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
'Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
'Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal param As Long) As Long
'Public Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As rect) As Long
'Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
'Public Declare PtrSafe Function DrawMenuBar Lib "user32.dll" (ByVal hwnd As Long) As Long
'Public Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'Public Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hmenu As Long, ByVal ideEnableItem As Long, ByVal enable As Long) As Integer
'Public Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hmenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Public Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hmenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Public Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As Long
'Public Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
''Mouse Functions
'Public Declare PtrSafe Function SetCursorPOs Lib "user32" Alias "SetCursorPos" (ByVal x As Long, ByVal y As Long) As Long
'Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'Public Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
'Public Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
'
''Other Functions
'Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Public Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare PtrSafe Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVallpDirectory As String, ByVal lpResult As String) As Long
'Public Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Public Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
'Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
'Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
''Convert the handle into an OLE IPicture interface.
'Public Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
''Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
'Public Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
''Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
'Public Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'#Else
'32 bit version of VBA
'Windows Functions
Private Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal lngHWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowExA Lib "user32" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal param As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As rect) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function DrawMenuBar Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hmenu As Long, ByVal ideEnableItem As Long, ByVal enable As Long) As Integer
Private Declare Function DeleteMenu Lib "user32" (ByVal hmenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hmenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, ByVal MullP As Long, ByVal NullP2 As Long, lpLastWriteTime As FILETIME) As Long

Private Declare Function SetFileTimeCreate Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, CreateTime As FILETIME, ByVal LastAccessTime As Long, ByVal LastModified As Long) As Long
Private Declare Function SetFileTimeLastAccess Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, ByVal CreateTime As Long, LastAccessTime As FILETIME, ByVal LastModified As Long) As Long
Private Declare Function SetFileTimeLastModified Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, ByVal CreateTime As Long, ByVal LastAccessTime As Long, LastModified As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Mouse Functions
Public Declare Function SetCursorPOs Lib "user32" Alias "SetCursorPos" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)

'Other Functions
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVallpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
'Convert the handle into an OLE IPicture interface.
Public Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
Public Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal N1 As Long, ByVal N2 As Long, ByVal un2 As Long) As Long
''#End If

'Program run function
Public Function run(ByVal msg As String) As Boolean
Debug.Print msg
run = True
End Function

'Debug Function (Func: 0 - Function, 1 - Subprogram, 2 - Nothing)
Public Function debug_err(Optional ByVal name As String, Optional ByVal derr As String, Optional ByVal merr As String, Optional ByVal Func As Byte, Optional suppress As Boolean, Optional msgbx As Boolean) As String
Dim i As Long
Dim ERR As String
Dim msg As New Collection

'Message Lists
msg.add "fna: This Function is not available now!"
msg.add "wsnex: The worksheet does not exist."
If InStr(1, derr, "d") = 1 Then
    msg.add "d1: (Function debug_err error!! If name is null, derr and merr must be null as well.)"
    msg.add "d2: (Function debug_err error!! Either derr or merr should be null.)"
    msg.add "d3: (Function debug_err error!! No such derr is found!!)"
ElseIf InStr(1, derr, "n") = 1 Then
    msg.add "nnum: The input is not numeric."
        If Left(derr, 4) = "nnum" And InStr(1, derr, "-") <> 0 Then
            msg.add derr & ": " & Trim(mid(derr, InStr(1, derr, "-") + 1, Len(derr) - InStr(1, derr, "-"))) & " is not numeric."
        End If
    msg.add "npi: The input are not positive integers."
        If Left(derr, 3) = "npis" And InStr(1, derr, "-") <> 0 Then
            msg.add derr & ": " & Trim(mid(derr, InStr(1, derr, "-") + 1, Len(derr) - InStr(1, derr, "-"))) & " are not positive integers."
        End If
    msg.add "npi: The input is not positive integer."
        If Left(derr, 3) = "npi" And InStr(1, derr, "-") <> 0 Then
            msg.add derr & ": " & Trim(mid(derr, InStr(1, derr, "-") + 1, Len(derr) - InStr(1, derr, "-"))) & " is not positive integer."
        End If
    msg.add "nprob: The input is not probability."
        If Left(derr, 5) = "nprob" And InStr(1, derr, "-") <> 0 Then
            msg.add derr & ": " & Trim(mid(derr, InStr(1, derr, "-") + 1, Len(derr) - InStr(1, derr, "-"))) & " is not probability."
        End If
    msg.add "nay: The array input is not array."
        If Left(derr, 3) = "nay" And InStr(1, derr, "-") <> 0 Then
            msg.add derr & ": " & Trim(mid(derr, InStr(1, derr, "-") + 1, Len(derr) - InStr(1, derr, "-"))) & " is not array."
        End If
    msg.add "nmx: The array input is not a matrix"
        If Left(derr, 3) = "nmx" And InStr(1, derr, "-") <> 0 Then
            msg.add derr & ": " & Trim(mid(derr, InStr(1, derr, "-") + 1, Len(derr) - InStr(1, derr, "-"))) & " is not a matrix."
        End If
    msg.add "njay: The array input is not a jagged array."
    msg.add "nnumay: The array input is not numeric."
    msg.add "nmx: The array input is not a matrix"
    msg.add "nsqmx: The array input is not a square matrix."
End If
If IsNumeric(Left(derr, 1)) = True Then
    i = Left(derr, 1)
    msg.add i & "d: The array input is not " & i & " dimension."
    If Left(derr, 2) = i & "d" And InStr(1, derr, "-") <> 0 Then
        msg.add derr & ": " & Trim(mid(derr, InStr(1, derr, "-") + 1, Len(derr) - InStr(1, derr, "-"))) & " is not a " & i & " - dimensional array."
    End If
    msg.add i & "du: The array input cannot be " & (i + 1) & "- dimension or above."
    If Left(derr, 2) = i & "du" And InStr(1, derr, "-") <> 0 Then
        msg.add derr & ": " & Trim(mid(derr, InStr(1, derr, "-") + 1, Len(derr) - InStr(1, derr, "-"))) & " cannot be " & (i + 1) & " - dimension or above."
    End If
End If

'Start
If Func = 0 Then
    ERR = "Function error!!"
ElseIf Func = 1 Then
    ERR = "Subprogram error!!"
ElseIf Func = 2 Then
    ERR = "Error!!"
End If
If name = "" Then
    If derr = "" And merr = "" Then
        ERR = ERR
    ElseIf derr <> "" Or merr <> "" Then
        debug_err name, "d1"
        Exit Function
    End If
ElseIf name <> "" Then
    If Func = 0 Then
        ERR = Left(ERR, 8) & " " & name & Right(ERR, 8)
    ElseIf Func = 1 Then
        ERR = Left(ERR, 10) & " " & name & Right(ERR, 8)
    ElseIf Func = 2 Then
        ERR = name & ERR
    End If
    If derr = "" And merr = "" Then
        ERR = ERR
    ElseIf derr = "" And merr <> "" Then
        ERR = ERR & " " & merr
    ElseIf derr <> "" And merr <> "" Then
        debug_err name, "d2"
        Exit Function
    ElseIf derr <> "" And merr = "" Then
        For i = 1 To msg.Count
            If derr = Left(msg.Item(i), (InStr(1, msg.Item(i), ":") - 1)) Then
                ERR = ERR & " " & mid(msg.Item(i), (InStr(1, msg.Item(i), ":") + 2), Len(msg.Item(i)) - (InStr(1, msg.Item(i), ":") + 1))
                GoTo ptmsg
            End If
        Next i
        debug_err name, "d3"
        Exit Function
    End If
End If
ptmsg:
If suppress = False Then
    If msgbx = False Then
        Debug.Print ERR
    Else
        MsgBox ERR, , "Debug Error!!"
    End If
End If
debug_err = ERR
End Function

'Can debug print or not (is the type of the array strings?)
Public Function isstr(ByVal v As Variant) As Boolean
Dim A As Variant, b As Variant
On Error Resume Next
If IsArray(v) = True Then
    For Each A In v
        b = CStr(A)
        If ERR.Number > 0 Then
            isstr = False
            Exit Function
        End If
    Next A
    isstr = True
Else
    b = CStr(v)
    If ERR.Number > 0 Then
        isstr = False
    Else
        isstr = True
    End If
End If
End Function

'Debug print values (up to 4 dimension array)
Public Function prt(ByVal v As Variant, Optional nck As Boolean) As Variant
Dim i As Long, j As Long, k As Long, l As Long
Dim R As Variant
Dim str As String, str1 As String
Dim d As Integer
If nck = False Then
    If isstr(v) = False Then
        debug_err "prt", , "The input cannot be converted to Strings to be printed."
        Exit Function
    End If
End If
If IsArray(v) = False Then
    Debug.Print v
Else
    R = ayrge(v)
    d = UBound(R, 2)
    If d = 1 Then
        For i = R(1, 1) To R(2, 1)
            If i = R(1, 1) Then
                str = CStr(v(i))
            Else
                str = str & ", " & CStr(v(i))
            End If
        Next i
        Debug.Print str
    ElseIf d = 2 Then
        For i = R(1, 1) To R(2, 1)
            str = ""
            For j = R(1, 2) To R(2, 2)
                If j = R(1, 2) Then
                    str = CStr(v(i, j))
                Else
                    str = str & ", " & CStr(v(i, j))
                End If
            Next j
            Debug.Print str
        Next i
    ElseIf d = 3 Then
        For k = R(1, 1) To R(2, 1)
            For i = R(1, 2) To R(2, 2)
                str = ""
                For j = R(1, 3) To R(2, 3)
                    If j = R(1, 3) Then
                        str = CStr(v(k, i, j))
                    Else
                        str = str & ", " & CStr(v(k, i, j))
                    End If
                Next j
                Debug.Print str
            Next i
            Debug.Print ""
        Next k
    ElseIf d = 4 Then
        For k = R(1, 1) To R(2, 1)
            For i = R(1, 3) To R(2, 3)
                str = ""
                For l = R(1, 2) To R(2, 2)
                    str1 = ""
                    For j = R(1, 4) To R(2, 4)
                        If j = R(1, 4) Then
                            str1 = CStr(v(k, l, i, j))
                        Else
                            str1 = str1 & ", " & CStr(v(k, l, i, j))
                        End If
                    Next j
                    If l = R(1, 2) Then
                        str = str1
                    Else
                        str = str & " * " & str1
                    End If
                Next l
                Debug.Print str
            Next i
            Debug.Print ""
        Next k
    Else
        debug_err "prt", "4du"
    End If
End If
End Function

'Run Time Analysis of Algorithms
Public Function AOAS(ByVal sub1 As String, ByVal sub2 As String, Optional ByVal n As Long) As Variant
Dim t(1 To 3) As Double
Dim t1 As Double, t2 As Double, S As Double, R As Double
Dim F(1 To 4, 1 To 1) As String
If n = 0 Then n = 1
t(1) = Timer()
For i = 1 To n
    Application.run sub1
Next i
t(2) = Timer()
For i = 1 To n
    Application.run sub2
Next i
t(3) = Timer()
t1 = (t(2) - t(1)) / n
t2 = (t(3) - t(2)) / n
S = Abs(t1 - t2)
R = t1 / t2
F(1, 1) = "The time program1 takes " & t1 & " seconds."
F(2, 1) = "The time program2 takes " & t2 & " seconds."
F(3, 1) = "The time difference is " & S & " seconds."
F(4, 1) = "The time ratio is " & R & "."
AOAS = F
End Function

'Run Time Analysis of Algorithms
Public Function AOA1(ByVal sub1 As String, ByVal sub2 As String, Optional ByVal n As Long) As Variant
Dim t(1 To 5) As Double
Dim t1 As Double, t2 As Double, S As Double, R As Double
Dim F(1 To 4, 1 To 1) As String
Dim NS(1 To 2) As Long
If n = 0 Then n = 1
NS(1) = Int(n / 2)
NS(2) = n - NS(1)
t(1) = Timer()
For i = 1 To NS(1)
    Application.run sub1
Next i
t(2) = Timer()
For i = 1 To NS(2)
    Application.run sub2
Next i
t(3) = Timer()
For i = 1 To NS(1)
    Application.run sub2
Next i
t(4) = Timer()
For i = 1 To NS(2)
    Application.run sub1
Next i
t(5) = Timer()
t1 = (t(2) - t(1) + t(5) - t(4)) / n
t2 = (t(3) - t(2) + t(4) - t(3)) / n
S = Abs(t1 - t2)
R = sigfig(t1 / t2, 3)
F(1, 1) = "The time program1 takes " & t1 & " seconds."
F(2, 1) = "The time program2 takes " & t2 & " seconds."
F(3, 1) = "The time difference is " & S & " seconds."
F(4, 1) = "The time ratio is " & R & "."
AOA1 = F
End Function

'Run Time Analysis of Algorithms
Public Function AOA2(ByVal sub1 As String, ByVal sub2 As String, Optional ByVal n As Long) As Variant
Dim t(1 To 9) As Double
Dim t1 As Double, t2 As Double, S As Double, R As Double
Dim F(1 To 4, 1 To 1) As String
Dim NS(1 To 4) As Long
If n = 0 Then n = 1
NS(1) = Int(n / 4)
NS(2) = NS(1)
NS(3) = NS(1)
NS(4) = n - NS(1) * 3
t(1) = Timer()
For i = 1 To NS(1)
    Application.run sub1
Next i
t(2) = Timer()
For i = 1 To NS(1)
    Application.run sub2
Next i
t(3) = Timer()
For i = 1 To NS(2)
    Application.run sub2
Next i
t(4) = Timer()
For i = 1 To NS(2)
    Application.run sub1
Next i
t(5) = Timer()
For i = 1 To NS(3)
    Application.run sub1
Next i
t(6) = Timer()
For i = 1 To NS(3)
    Application.run sub2
Next i
t(7) = Timer()
For i = 1 To NS(4)
    Application.run sub2
Next i
t(8) = Timer()
For i = 1 To NS(4)
    Application.run sub1
Next i
t(9) = Timer()
t1 = (t(2) - t(1) + t(5) - t(4) + t(6) - t(5) + t(9) - t(8)) / n
t2 = (t(3) - t(2) + t(4) - t(3) + t(7) - t(6) + t(8) - t(7)) / n
S = Abs(t1 - t2)
R = t1 / t2
F(1, 1) = "The time program1 takes " & t1 & " seconds."
F(2, 1) = "The time program2 takes " & t2 & " seconds."
F(3, 1) = "The time difference is " & S & " seconds."
F(4, 1) = "The time ratio is " & R & "."
AOA2 = F
End Function

'Run Time Analysis of Algorithms
Public Function AOAF(ByVal sub1 As String, ByVal sub2 As String, Optional ByVal n As Long) As Variant
Dim i As Long, j As Long, k As Long
Dim t() As Double
Dim tt() As Double
Dim d As Long, d1
Dim NS() As Long
Dim t1 As Double, t2 As Double, TS As Double, R As Double
Dim F(1 To 4, 1 To 1) As String
Dim S(1 To 2) As String
If n = 0 Then n = 1
S(1) = sub1
S(2) = sub2
d = n ^ 0.5
ReDim NS(1 To d) As Long
ReDim t(1 To d * 2 + 1) As Double
ReDim tt(1 To d * 2) As Double
d1 = Int(n / d)
For i = 1 To (d - 1)
    NS(i) = d1
Next i
NS(d) = n - d1 * (d - 1)
'Start
k = 1
For i = 1 To d
    If i Mod 2 = 1 Then
        t(k) = Timer()
        For j = 1 To NS(i)
            Application.run S(1)
        Next j
        k = k + 1
        t(k) = Timer()
        For j = 1 To NS(i)
            Application.run S(2)
        Next j
        k = k + 1
    ElseIf i Mod 2 = 0 Then
        t(k) = Timer()
        For j = 1 To NS(i)
            Application.run S(2)
        Next j
        k = k + 1
        t(k) = Timer()
        For j = 1 To NS(i)
            Application.run S(1)
        Next j
        k = k + 1
    End If
    If i = d Then t(k) = Timer()
Next i
For i = 1 To d * 2
    tt(i) = t(i + 1) - t(i)
Next i
For i = 1 To d * 2
    If i Mod 4 = 0 Or i Mod 4 = 1 Then
        t1 = t1 + tt(i)
    ElseIf i Mod 4 = 2 Or i Mod 4 = 3 Then
        t2 = t2 + tt(i)
    End If
Next i
t1 = t1 / n
t2 = t2 / n
TS = Abs(t1 - t2)
R = sigfig(t1 / t2, 3)
F(1, 1) = "The time program1 takes " & t1 & " seconds."
F(2, 1) = "The time program2 takes " & t2 & " seconds."
F(3, 1) = "The time difference is " & TS & " seconds."
F(4, 1) = "The time ratio is " & R & "."
AOAF = F
End Function

'Run Time Analysis of Algorithms
Public Function AOA(ByVal sub1 As String, ByVal sub2 As String, Optional ByVal n As Long) As Variant
AOA = AOA1(sub1, sub2, n)
End Function

'******************************************************************************************************************************************************
'B!: General Functions

'Boolean And (More than 2 boolean variables)
Public Function booland(ByVal b As Variant) As Boolean
Dim v As Variant
If IsArray(b) = True Then
    For Each v In b
        If v = False Then
            booland = False
            Exit Function
        End If
    Next v
    booland = True
Else
    debug_err "nay"
End If
End Function

'Boolean Or (More than 2 boolean variables)
Public Function boolor(ByVal b As Variant) As Boolean
Dim v As Variant
If IsArray(b) = True Then
    For Each v In b
        If v = True Then
            boolor = True
            Exit Function
        End If
    Next v
    boolor = False
Else
    debug_err "nay"
End If
End Function

'B1!: Maths Functions
'isint (options: P (Positive) or N (Negative), O (odd) or E (even))
Public Function isint(ByVal n As Variant, Optional ByVal sign As String, Optional oddeven As String) As Boolean
Dim b As Boolean
b = IsNumeric(n)
If b = False Then
    isint = False
    Exit Function
End If
b = (CLng(n) = n)
If b = False Then
    isint = False
    Exit Function
End If
If sign = "" Then
ElseIf sign = "P" Then
    If Abs(n) <> n Or n = 0 Then
        isint = False
        Exit Function
    End If
ElseIf sign = "N" Then
    If Abs(n) = n Or n = 0 Then
        isint = False
        Exit Function
    End If
Else
    debug_err "isint", , "Please specify correct sign (P (for positive) or N (for negative))."
End If
If oddeven = "" Then
ElseIf oddeven = "O" Then
    If isint(n / 2) = True Then
        isint = False
        Exit Function
    End If
ElseIf oddeven = "E" Then
    If isint(n / 2) = False Then
        isint = False
        Exit Function
    End If
Else
    debug_err "isint", , "Please specify correct oddeven indicator (O (for odd) or E (for even))."
End If
isint = True
End Function

'is probability?
Public Function isprob(ByVal P As Variant) As Boolean
If IsNumeric(P) = True Then
    If P >= 0 And P <= 1 Then
        isprob = True
    Else
        isprob = False
    End If
Else
    isprob = False
End If
End Function

'is percent?
Public Function ispercent(ByVal P As Variant) As Boolean
If IsNumeric(P) = True Then
    If P >= 0 And P <= 100 Then
        ispercent = True
    Else
        ispercent = False
    End If
Else
    ispercent = False
End If
End Function

'is prime?
Public Function isprime(ByVal n As Variant) As Boolean
Dim i As Long
Dim t As Long
If isint(n, "P") = True Then
    t = Ceil(n ^ 0.5)
    If n = 1 Then
        isprime = False
    ElseIf n < 4 Then
        isprime = True
    ElseIf n Mod 2 = 0 Then
        isprime = False
    ElseIf n < 9 Then
        isprime = True
    ElseIf n Mod 3 = 0 Then
        isprime = False
    Else
        i = 5
        Do While i <= t
            If n Mod i = 0 Then
                isprime = False
                Exit Function
            End If
            If n Mod (i + 2) = 0 Then
                isprime = False
                Exit Function
            End If
            i = i + 6
        Loop
        isprime = True
    End If
Else
    debug_err "isprime", "npi - n"
End If
End Function

'Another algorithm for isprime
Public Function isprime_a(ByVal n As Variant) As Boolean

End Function

'Generating nth Prime numbers (all: generating n Prime numbers)
Public Function prime(ByVal n As Long, Optional ByVal all As Boolean) As Variant
Dim i As Long, j As Long
Dim Q As Long
Dim P() As Long
ReDim P(1 To n) As Long
P(1) = 1
P(2) = 2
i = 3
Q = i
Do Until P(n) <> 0
    For j = 2 To n
        If (i ^ 0.5) < P(j) Then GoTo fd
        If i Mod P(j) = 0 Then
            GoTo SKip
        End If
    Next j
fd:
    P(Q) = i
    Q = Q + 1
SKip:
    i = i + 1
Loop
If all = False Then
    prime = P(n)
ElseIf all = True Then
    prime = P
End If
End Function

'Product of n prime numbers
Public Function primeproduct(ByVal n As Long) As Long
Dim A As Variant
Dim P As Long: P = 1
A = prime(n, True)
For i = 1 To n
    P = P * A(i)
Next i
primeproduct = P
End Function

'Convert to Binary number
Public Function binary(ByVal n As Long) As String
Dim b As String
Dim i As Long
i = 1
Do While (n > 0)
    If ((n Mod (2 ^ i)) <> 0) Then
        b = "1" & b
        n = n - (2 ^ (i - 1))
    Else
        b = "0" & b
    End If
    i = i + 1
Loop
binary = b
End Function

'Reverse of the number (option: b bas base (default base 10))
Public Function renum(ByVal n As Long, Optional ByVal b As Long) As Long
Dim R As Long
If b = 0 Then b = 10
R = 0
Do While n >= 1
    R = b * R + n Mod b
    n = Int(n / b)
Loop
renum = R
End Function

'Is Palindrome number
Public Function ispalindrome(ByVal n As Long) As Boolean
If n = renum(n) Then
    ispalindrome = True
End If
End Function

'Out of Pre-set range
Public Function outrange(ByVal n As Variant, ByVal low As Variant, ByVal up As Variant) As Boolean
If IsNumeric(n) = False Then
    debug_err "Outrange", "nnum"
    Exit Function
End If
If IsNumeric(low) = False Then
    debug_err "Outrange", "nnum"
    Exit Function
End If
If IsNumeric(up) = False Then
    debug_err "Outrange", "nnum"
    Exit Function
End If
If n < low Or n > up Then
    debug_err "Outrange", , "The number - " & n & " is out of the Pre-set range (" & low & " - " & up & ")."
    Exit Function
End If
outrange = True
End Function

'Pi
Public Function Pi() As Double
Pi = 4 * Atn(1)
End Function

'Log (general case, specify the base)
Public Function logB(ByVal n As Double, Optional ByVal b As Double) As Double
If b = 0 Then
    logB = Log(n)
ElseIf b > 0 And b <> 1 Then
    logB = Log(n) / Log(b)
Else
    debug_err "logB", , "B must be positive and not equal to 1."
End If
End Function

'Factorial (Non-Resursive Algorithm)
Public Function Fact(ByVal n As Long) As Long
If n = 0 Then
    Fact = 1
ElseIf n > 0 Then
    Fact = 1
    For i = 1 To n
        Fact = Fact * i
    Next i
ElseIf n < 0 Then
    debug_err "Fact", "npi - n"
End If
End Function

'Factorial (Recursive Algorithm)
Public Function Fact_r(ByVal n As Long) As Long
If n < 0 Then
    debug_err "Fact", "npi - n"
ElseIf n = 0 Then
    Fact_r = 1
Else
    Fact_r = n * Fact_r(n - 1)
End If
End Function

'Double Factorial
Public Function DFact(ByVal n As Long) As Long
DFact = 1
Do Until n < 1
    DFact = DFact * n
    n = n - 2
Loop
End Function

'Permutation
Public Function Npr(ByVal n As Long, R As Long) As Long
Dim NR As Long
If n >= 0 And R >= 0 And R <= n Then
NR = n - R
Npr = 1
    Do Until n = NR
        Npr = Npr * n
        n = n - 1
    Loop
Else
    debug_err "Npr", , "Wrong Specification of N or R, please check!!"
End If
End Function

'Combination
Public Function ncr(ByVal n As Long, R As Long) As Long
Dim i As Long, j As Long
Dim A As Variant
If n >= 0 And R >= 0 And R <= n Then
    ncr = ncrc(n, R)
Else
    debug_err "ncr", , "Wrong Specification of N or R, please check!!"
End If
End Function
    
    'Recursive algorithm
    Public Function ncra(ByVal n As Long, R As Long) As Long
        If R = 0 Or R = n Then
            ncra = 1
        Else
            ncra = ncra(n - 1, R - 1) + ncra(n - 1, R)
        End If
    End Function
    
    'Iteritive algorithm (Pascal Triangle Algorithm)
    Public Function ncrb(ByVal n As Long, R As Long) As Long
    Dim i As Long, j As Long
    Dim A() As Long
    ReDim A(0 To n, 0 To n) As Long
    For i = 0 To n
        A(i, 0) = 1
        A(i, i) = 1
    Next i
    For i = 2 To n
        For j = 1 To n - 1
            A(i, j) = A(i - 1, j - 1) + A(i - 1, j)
        Next j
    Next i
    ncrb = A(n, R)
    End Function
    
    'Iteritive algorithm (n! / r!(n-r)!)
    Public Function ncrc(ByVal n As Long, ByVal R As Long) As Long
    Dim i As Long, j As Long
    Dim G As Long
    Dim NA As Variant
    Dim RA As Variant
    Dim NS As Long
    Dim RS As Long
    If (R = 0) Or (R = n) Then
        ncrc = 1
        Exit Function
    ElseIf R = 1 Then
        ncrc = n
        Exit Function
    ElseIf R > 1 And R < n Then
        If R > (n - R) Then
            R = n - R
        End If
    
    End If
    ReDim NA(1 To R) As Variant
    ReDim RA(1 To R) As Variant
    For i = 1 To R
        RA(i) = i
        NA(i) = n - i + 1
    Next i
    For i = R To 1 Step -1
        For j = 1 To R
            If (NA(j) Mod RA(i)) = 0 Then
                NA(j) = NA(j) / RA(i)
                RA(i) = 1
            End If
        Next j
    Next i
    For i = 1 To R
        Do While (RA(i) > 1)
            For j = 1 To R
                G = GCD_b(NA(j), RA(i))
                If G > 1 Then
                    NA(j) = NA(j) / G
                    RA(i) = RA(i) / G
                End If
            Next j
        Loop
    Next i
    NS = 1
    RS = 1
    For i = 1 To R
        NS = NS * NA(i)
        RS = RS * RA(i)
    Next i
    ncrc = NS / RS
    End Function

'Multinominal (The ratio between factorial of the sum of the number and the products of the factorial of the number)
Public Function multinominal(ByVal NS As Variant) As Long
Dim i As Long
Dim S As Long, fp As Long
Dim R As Variant
Dim A As Variant
If isay(NS) = False Then
    If isint(NS, "P") = True Then
        multinominal = 1
    Else
        debug_err "multinominal", "npis - n"
    End If
Else
    If isay(NS, , True, 1, True) = True Then
        R = ayrge(NS)
        A = NS
        fp = 1
        For i = R(1, 1) To R(2, 1)
            fp = fp * Fact(A(i))
        Next i
        S = Fact(aysum(A))
        multinominal = S / fp
    Else
        debug_err "multinominal", "1d - n"
    End If
End If
End Function

'Convert Radian to Degree
Public Function Deg(ByVal Rad As Double) As Double
Deg = (Rad / Pi) * 180
End Function

'Convert Degree to Radian
Public Function Rad(ByVal Deg As Double) As Double
Rad = Deg / 180 * Pi
End Function

'Add Degree symbol to the number
Public Function degadd(ByVal n As Variant) As String
If IsNumeric(n) = True Then
    degadd = CStr(n) & "¢X"
Else
    degadd = CStr(n)
End If
End Function

'Number of decimal places
Public Function dec(ByVal num As Variant) As Long
Dim n As Variant
Dim i As Long
n = num
Do Until n = Int(n)
    n = n * 10
    i = i + 1
Loop
dec = i
End Function

'Normal Rounding (VBA uses banker's Rounding)
Public Function RoundN(ByVal num As Variant, ByVal dc As Long) As Variant
Dim i As Long
num = num * (10 ^ dc)
i = Int(num)
If num - i >= 0.5 Then
    num = Int(num) + 1
ElseIf num - i < 0.5 Then
    num = Int(num)
End If
RoundN = num / (10 ^ dc)
End Function

'Rounding (Generalised version) 'rd: R = normal rounding, C = Ceiling, F = Floor, B = Banker's Rounding
Public Function RoundG(ByVal num As Variant, Optional ByVal dc As Long, Optional rd As String) As Variant
Dim A As Variant
If IsNumeric(num) = True Then
    If dc <= dec(num) Then
        A = num * (10 ^ dc)
            If rd = "" Then rd = "R"
            If rd = "R" Then
                A = RoundN(A, 0)
            ElseIf rd = "C" Then
                A = Int(A) + 1
            ElseIf rd = "F" Then
                A = Int(A)
            ElseIf rd = "B" Then
                A = Round(A)
            End If
        A = A / (10 ^ dc)
        RoundG = A
    Else
        debug_err "RoundG", , "dc greater than decimal places of the number."
    End If
Else
    debug_err "RoundG", "nnum"
End If
End Function

'Rounding to multiples
Public Function Mround(ByVal num As Variant, ByVal m As Long, Optional rd As String) As Long
Dim A As Variant
If IsNumeric(num) = True Then
    If isint(m, "P") = True Then
        A = num / m
        If rd = "" Then rd = "R"
        If rd = "R" Then
            A = RoundN(A, 0)
        ElseIf rd = "C" Then
            A = Int(A) + 1
        ElseIf rd = "F" Then
            A = Int(A)
        End If
        Mround = A * m
    Else
        debug_err "Mround", "np - M"
    End If
Else
    debug_err "Mround", "nnum"
End If
End Function

'Floor
Public Function Floor(ByVal num As Variant, Optional ByVal dc As Long) As Variant
Floor = RoundG(num, dc, "F")
End Function

'Ceiling
Public Function Ceil(ByVal num As Variant, Optional ByVal dc As Long) As Variant
Ceil = RoundG(num, dc, "C")
End Function

'Significant figures
Public Function sigfig(ByVal num As Variant, ByVal sig As Long, Optional rd As String) As Variant
Dim dc As Long
If IsNumeric(num) = True Then
    If sig > 0 Then
        dc = sig - (Int(logB(num, 10)) + 1)
        If rd = "" Then rd = "R"
        If rd = "R" Then
            sigfig = RoundG(num, dc, "R")
        ElseIf rd = "C" Then
            sigfig = RoundG(num, dc, "C")
        ElseIf rd = "F" Then
            sigfig = RoundG(num, dc, "F")
        Else
            debug_err "sigfig", , "Please specify the correct type of rounding method."
        End If
    Else
        debug_err "sigfig", "npi - sig"
    End If
Else
    debug_err "sigfig", "nnum"
End If
End Function

'Greatest Common Divisor
Public Function GCD(ByVal NS As Variant) As Variant
Dim i As Long, k As Long: k = 2
Dim m As Long
Dim P As Long
Dim v As Long: v = 1
Dim R As Variant
If isay(NS, , True, 1, True, "P") = True And ayct(NS) >= 2 Then
    m = aymin(NS)
    R = ayrge(NS)
    Do
        P = prime(k)
        For i = R(1, 1) To R(2, 1)
            If isint(NS(i) / P) = False Then
                k = k + 1
                GoTo SKip
            End If
        Next i
        v = v * P
        For i = R(1, 1) To R(2, 1)
            NS(i) = NS(i) / P
        Next i
SKip:
    Loop While (m ^ 0.5) > P
    GCD = v
Else
    debug_err "GCD", , "Please input an array with 2 or more positive integers"
End If
End Function

'Greatest Common Divisor (two variables)
Public Function GCD_b(ByVal x As Long, ByVal y As Long) As Long
Dim z As Long
Do Until y = 0
    z = y
    y = x Mod y
    x = z
Loop
GCD_b = x
End Function

'Lowest common multiple
Public Function LCM(ByVal NS As Variant) As Variant
Dim i As Long, k As Long: k = 2
Dim m As Long
Dim P As Long
Dim v As Long: v = 1
Dim R As Variant
Dim c As Long: c = 2
If isay(NS, , True, 1, True, "P") = True And ayct(NS) >= 2 Then
    m = aymin(NS)
    R = ayrge(NS)
    Do
        P = prime(k)
        For i = R(1, 1) To R(2, 1)
            If isint(NS(i) / P) = False Then
                k = k + 1
                GoTo SKip
            End If
        Next i
        v = v * P
        For i = R(1, 1) To R(2, 1)
            NS(i) = NS(i) / P
        Next i
SKip:
    Loop While (m ^ 0.5) > P
    For i = R(1, 1) To R(2, 1)
        v = v * NS(i)
    Next i
    LCM = v
Else
    debug_err "LCM", , "Please input an array with 2 or more positive integers"
End If
End Function

'Greatest Common Divisor (two variables)
Public Function LCM_b(ByVal x As Long, ByVal y As Long) As Long
LCM_b = x * y / GCD_b(x, y)
End Function

'Long integer comparison
Function LComp(ByVal x As String, ByVal y As String) As Integer
Dim i As Long
Dim l(1 To 2) As Long
l(1) = Len(x)
l(2) = Len(y)
If l(1) > l(2) Then
    LComp = 1
ElseIf l(1) < l(2) Then
    LComp = -1
ElseIf l(1) = l(2) Then
    For i = 1 To l(1)
        If mid(x, i, 1) > mid(y, i, 1) Then
            LComp = 1
            Exit Function
        ElseIf mid(x, i, 1) < mid(y, i, 1) Then
            LComp = -1
            Exit Function
        End If
    Next i
    LComp = 0
End If
End Function

'Long integer addition
Function Ladd(ByVal x As String, ByVal y As String) As String
Dim i As Long, j As Long
Dim n As String
Dim A(1 To 2) As String
Dim d() As Integer
Dim l(0 To 2) As Integer
Dim NB(1 To 2) As Boolean
Dim c As Integer
A(1) = x
A(2) = y
For j = 1 To 2
    NB(j) = (InStr(1, A(j), "-") = 1)
    If NB(j) = False Then
        A(j) = A(j)
    Else
        A(j) = mid(A(j), 2, Len(A(j)) - 1)
    End If
Next j
If NB(1) = NB(2) Then
    l(1) = Len(A(1))
    l(2) = Len(A(2))
    If l(1) >= l(2) Then
        l(0) = l(1)
    Else
        l(0) = l(2)
    End If
    ReDim d(0 To 2, 1 To l(0) + 1) As Integer
    For i = 1 To (l(0) + 1)
        For j = 1 To 2
            If i <= l(j) Then
                d(j, i) = CInt(mid(A(j), l(j) - i + 1, 1))
            End If
        Next j
        d(0, i) = d(0, i) + d(1, i) + d(2, i)
        If d(0, i) >= 10 Then
            d(0, i) = d(0, i) - 10
            d(0, i + 1) = d(0, i + 1) + 1
        End If
        If i <= l(0) Then
            n = d(0, i) & n
        Else
            If d(0, i) > 0 Then
                n = d(0, i) & n
            End If
        End If
    Next i
    If NB(1) = True Then
        n = "-" & n
    End If
Else
    c = LComp(A(1), A(2))
    If c = 0 Then
        n = "0"
    Else
        n = Lsub(A(1), A(2))
        If c = 1 Then
            If NB(1) = True Then
                n = "-" & n
            End If
        ElseIf c = -1 Then
            If NB(1) = True Then
                n = mid(n, 2, Len(n) - 1)
            End If
        End If
    End If
End If
Ladd = n
End Function

'Long integer substraction
Function Lsub(ByVal x As String, ByVal y As String) As String
Dim i As Long, j As Long
Dim t As String
Dim n As String
Dim A(1 To 2) As String
Dim E() As Boolean
Dim d() As Integer
Dim m As Boolean
Dim l(0 To 2) As Integer
Dim NB(1 To 2) As Boolean
A(1) = x
A(2) = y
For j = 1 To 2
    NB(j) = (InStr(1, A(j), "-") = 1)
    If NB(j) = False Then
        A(j) = A(j)
    Else
        A(j) = mid(A(j), 2, Len(A(j)) - 1)
    End If
Next j
If NB(1) = False And NB(2) = False Then
    If LComp(A(2), A(1)) = 1 Then
        m = True
        t = A(1)
        A(1) = A(2)
        A(2) = t
    End If
    l(1) = Len(A(1))
    l(2) = Len(A(2))
    If l(1) >= l(2) Then
        l(0) = l(1)
    Else
        l(0) = l(2)
    End If
    ReDim d(0 To 2, 1 To l(0)) As Integer
    ReDim E(1 To l(0) + 1) As Boolean
    For i = 1 To l(1)
        d(1, i) = CInt(mid(A(1), l(1) - i + 1, 1))
    Next i
    For i = 1 To l(2)
        d(2, i) = CInt(mid(A(2), l(2) - i + 1, 1))
    Next i
    For i = 1 To l(0)
        d(0, i) = d(1, i) - d(2, i)
        E(i) = True
    Next i
    For i = l(0) To 1 Step -1
        If i = l(0) Then
            If d(0, i) = 0 Then
                E(i) = False
            End If
        Else
            If d(0, i) = 0 And E(i + 1) = False Then
                E(i) = False
            End If
        End If
    Next i
    For i = 1 To l(0)
        If d(0, i) < 0 Then
            d(0, i) = d(0, i) + 10
            d(0, i + 1) = d(0, i + 1) - 1
            If d(0, i + 1) = 0 Then
                E(i + 1) = False
            End If
        End If
    Next i
    For i = 1 To l(0)
        If (E(i) = False And i > 1) Then Exit For
        n = d(0, i) & n
    Next i
    If m = True Then
        n = "-" & n
    End If
ElseIf NB(1) = True And NB(2) = False Then
    n = "-" & Ladd(A(1), A(2))
ElseIf NB(1) = False And NB(2) = True Then
    n = Ladd(A(1), A(2))
Else
    If A(2) > A(1) Then
        n = Lsub(A(2), A(1))
    Else
        n = "-" & Lsub(A(1), A(2))
    End If
End If
Lsub = n
End Function

'Long integer multiplication
Function Lmulti(ByVal x As String, ByVal y As String) As String
Dim i As Long, j As Long, k As Long
Dim n As String
Dim A(1 To 2) As String
Dim d() As Integer
Dim E() As Boolean
Dim m As Long
Dim l(0 To 2) As Integer
Dim NB(1 To 2) As Boolean
NB(1) = (InStr(1, x, "-") = 1)
NB(2) = (InStr(1, y, "-") = 1)
If NB(1) = False Then
    A(1) = x
Else
    A(1) = mid(x, 2, Len(x) - 1)
End If
If NB(2) = False Then
    A(2) = y
Else
    A(2) = mid(y, 2, Len(y) - 1)
End If
l(1) = Len(A(1))
l(2) = Len(A(2))
If l(1) >= l(2) Then
    l(0) = l(1)
Else
    l(0) = l(2)
End If
ReDim d(0 To 2, 1 To (l(0) * 2)) As Integer
ReDim E(1 To l(0) * 2 + 1) As Boolean
For i = 1 To l(1)
    d(1, i) = CInt(mid(A(1), l(1) - i + 1, 1))
Next i
For i = 1 To l(2)
    d(2, i) = CInt(mid(A(2), l(2) - i + 1, 1))
Next i
For i = 1 To l(0)
    For j = 1 To l(0)
        d(0, i + j - 1) = d(0, i + j - 1) + d(1, i) * d(2, j)
        If d(0, i + j - 1) <> 0 Then
            E(i + j - 1) = True
        End If
    Next j
Next i
For i = 1 To (l(0) * 2)
    m = d(0, i)
    If d(0, i) >= 10 Then
        For j = 0 To 2
            If d(0, i) >= (10 ^ (3 - j)) Then
                d(0, i + 3 - j) = d(0, i + 3 - j) + Int(m / (10 ^ (3 - j)))
                E(i + 3 - j) = True
                m = m Mod (10 ^ (3 - j))
            End If
        Next j
        d(0, i) = m Mod 10
    End If
Next i
For i = 1 To (l(0) * 2)
    If E(i) = False Then Exit For
    n = d(0, i) & n
Next i
If NB(1) <> NB(2) Then
    n = "-" & n
End If
Lmulti = n
End Function

'Long integer division (d: decimal places) (for integer only)
Function Ldiv(ByVal x As String, ByVal y As String, Optional ByVal dc As Long) As String
Dim i As Long, j As Long, k As Long
Dim P As Long
Dim n As String
Dim A(1 To 2) As String
Dim E() As Boolean
Dim d() As Long
Dim m As Boolean
Dim l(0 To 2) As Long
Dim NB(1 To 2) As Boolean
Dim R As String
Dim DA As String
Dim dd As Long
A(1) = x
A(2) = y
l(1) = Len(A(1))
l(2) = Len(A(2))
If l(1) >= l(2) Then
    l(0) = l(1)
Else
    l(0) = l(2)
End If
ReDim d(0 To 2, 1 To l(0)) As Long
ReDim E(1 To l(0)) As Boolean
For i = 1 To l(1)
    d(1, i) = CLng(Trim(mid(A(1), l(1) - i + 1, 1)))
Next i
For i = 1 To l(2)
    d(2, i) = CLng(Trim(mid(A(2), l(2) - i + 1, 1)))
    R = R & CLng(mid(A(1), i, 1))
Next i
P = CInt(Trim(l(1) - l(2) + 1))
For i = 1 To P
    E(i) = True
Next i
Do Until (R = "0" Or (P = 1 And R < y))
    If LComp(R, y) = -1 Then
        If d(0, P) = 0 Then
            E(P) = False
        End If
        P = P - 1
        R = R & CStr(mid(A(1), (l(1) - P + 1), 1))
    End If
    DA = ""
    For i = 1 To 9
        DA = Ladd(DA, y)
        If LComp(DA, R) = 1 Then
            dd = i - 1
            Exit For
        ElseIf LComp(DA, R) = 0 Then
            dd = i
            Exit For
        End If
    Next i
    d(0, P) = Trim(CStr(dd))
    If LComp(DA, R) = 1 Then
        DA = Lsub(DA, y)
    End If
    R = Lsub(R, DA)
    If (R = 0 And P > 1) Then
        P = P - 1
        R = CStr(mid(A(1), (l(1) - P + 1), 1))
    End If
Loop
For i = 1 To l(0)
    If E(i) = False Then Exit For
    n = d(0, i) & n
Next i
Ldiv = n
End Function

'******************************************************************************************************************************************************
'B2!: Statistics Functions

'Error Function
Function Err_func(ByVal x As Variant) As Double
Dim n As Long, m As Long
Dim P As Double, Q As Double, pp As Double
n = 1
If x > 5 Then
    Err_func = 1
ElseIf x < -5 Then
    Err_func = -1
Else
    Do Until P = pp And n > 1
        pp = P
        Q = 1
        For m = 1 To n
            Q = Q / m * x ^ 2
        Next m
        P = P + Q / ((2 * n + 1) * (-1) ^ n) * x
        n = n + 1
    Loop
    Err_func = (P + x) * (2 / (Pi()) ^ 0.5)
End If
End Function

'Binomial Distribution
Public Function Dist_Bino(ByVal P As Double, ByVal n As Long, ByVal k As Long, Optional ByVal cum As Boolean) As Double
Dim i As Long
If isint(n, "P") = False Then
    debug_err "Dist_bino", "npi - n"
    Exit Function
End If
If isint(k, "P") = False Then
    debug_err "Dist_bino", "npi - k"
    Exit Function
End If
If isprob(P) = False Then
    debug_err "Dist_bino", "nprob - p"
    Exit Function
End If
If cum = False Then
    Dist_Bino = ncr(n, k) * P ^ k * (1 - P) ^ (n - k)
ElseIf cum = True Then
    For i = 0 To k
        Dist_Bino = Dist_Bino + ncr(n, i) * P ^ i * (1 - P) ^ (n - i)
    Next i
End If
End Function

'Poisson Distribution
Public Function Dist_Poisson(ByVal l As Double, ByVal k As Long, Optional ByVal cum As Boolean) As Double
Dim i As Long
If l < 0 Then
    debug_err "Dist_Poisson", , "L must be positive."
    Exit Function
End If
If isint(k, "P") = False Then
    debug_err "Dist_Poisson", , "k must be a positive integer"
    Exit Function
End If
If cum = False Then
    Dist_Poisson = Exp(1) ^ (-l) * (l ^ k) / Fact(k)
Else
    For i = 0 To k
        Dist_Poisson = Dist_Poisson + (l ^ i) / Fact(i)
    Next i
    Dist_Poisson = Dist_Poisson * Exp(1) ^ (-l)
End If
End Function

'Geometric Distribution (kth trial is the first success)
Public Function Dist_geometric(ByVal P As Double, ByVal k As Long, Optional ByVal cum As Boolean) As Double
If isint(k, "P") = False Then
    debug_err "Dist_geometric", , "k must be a positive integer"
    Exit Function
End If
If isprob(P) = False Then
    debug_err "Dist_geometric", , "p is not a probability, please check."
    Exit Function
End If
If cum = False Then
    Dist_geometric = (1 - P) ^ (k - 1) * P
ElseIf cum = True Then
    Dist_geometric = 1 - (1 - P) ^ k
End If
End Function

'Negative Binomial Distribution (number of success before r failures)
Public Function Dist_negbino(ByVal P As Double, ByVal R As Long, ByVal k As Long, Optional ByVal cum As Boolean) As Double
Dim i As Long
If isprob(P) = False Then
    debug_err "Dist_negbino", , "p is not a probability, please check."
    Exit Function
End If
If isint(k, "P") = False Then
    debug_err "Dist_negbino", , "k must be a positive integer"
    Exit Function
End If
If isint(R, "P") = False Then
    debug_err "Dist_negbino", , "r must be a positive integer"
    Exit Function
End If
If cum = False Then
    Dist_negbino = ncr((R + k - 1), k) * (1 - P) ^ R * P ^ k
ElseIf cum = True Then
    For i = 0 To k
        Dist_negbino = Dist_negbino + ncr((R + i - 1), i) * (1 - P) ^ R * P ^ i
    Next i
End If
End Function

'Hypergeometric Distribution
Public Function Dist_Hypergeometric(ByVal n As Long, ByVal k As Long, NS As Long, ks As Long, Optional ByVal cum As Boolean) As Double
Dim i As Long
If isint(n, "P") = False Then
    debug_err "Dist_negbino", , "N must be a positive integer"
    Exit Function
End If
If isint(k, "P") = False Then
    debug_err "Dist_negbino", , "K must be a positive integer"
    Exit Function
End If
If isint(NS, "P") = False Then
    debug_err "Dist_negbino", , "ns must be a positive integer"
    Exit Function
End If
If isint(ks, "P") = False Then
    debug_err "Dist_negbino", , "Ks must be a positive integer"
    Exit Function
End If
If NS <= n And ks <= k And k <= n Then
    If cum = False Then
        Dist_Hypergeometric = (ncr(k, ks) * ncr((n - k), (NS - ks))) / ncr(n, NS)
    ElseIf cum = True Then
        For i = 0 To ks
            Dist_Hypergeometric = Dist_Hypergeometric + (ncr(k, i) * ncr((n - k), (NS - i))) / ncr(n, NS)
        Next i
    End If
Else
    debug_err "Dist_Hypergeometric", , "parameter input are incorrect, please check."
End If
End Function

'Normal Distribution
Public Function Dist_normal(ByVal u As Double, ByVal S As Double, ByVal x As Double, Optional cum As Boolean) As Double
If cum = False Then
    Dist_normal = (1 / ((2 * Pi() * (S ^ 2)) ^ 0.5)) * (Exp((-(x - u) ^ 2) / (2 * (S ^ 2))))
ElseIf cum = True Then
    Dist_normal = 0.5 * (1 + Err_func((x - u) / (S * (2 ^ 0.5))))
End If
End Function

't distribution
Public Function Dist_t(ByVal d As Long, ByVal x As Double, Optional cum As Boolean) As Double
Dim i As Long
Dim F As Double
If d > 0 Then
    If cum = False Then
        F = 1
        i = d
        Do Until (i - 2) < 1
            F = F * (i - 1) / (i - 2)
            i = i - 2
        Loop
        If isint(d, "P", 1) = True Then
            F = (F / Pi()) / (d ^ 0.5)
        ElseIf isint(d, "P", 2) = True Then
            F = (F / 2) / (d ^ 0.5)
        End If
        F = F * ((1 + (x ^ 2 / d)) ^ (-(d + 1) / 2))
        Dist_t = F
    ElseIf cum = True Then
        
    End If
Else
    debug_err "Dist_t", , "d must be positive integer."
End If
End Function
'******************************************************************************************************************************************************

'B3!: String Functions
'Swap two values
Public Function swap(A As Variant, b As Variant) As Boolean
Dim t As Variant
t = A
A = b
b = t
swap = True
End Function

'Trim null value at the end
Public Function TrimNull(ByVal str As String) As String
    'Remove unnecessary null terminator at the end
    If InStrRev(str, chr$(0)) Then
        TrimNull = Left$(str, nullPos - 1)
    Else
        TrimNull = str
    End If
End Function

'Trim any sepecified value
Public Function TrimG(ByVal str As String, ByVal chr As Variant) As String
Dim b As Boolean
Dim i As Long, j As Long
Dim d As Long
Dim P(1 To 2) As Long
Dim R As Variant
Dim ln As Long
If isay(chr) = False Then
    ln = Len(str)
    For i = 1 To ln
        If mid(str, i, 1) <> chr Then
            P(1) = i
            Exit For
        End If
    Next i
    For i = 1 To ln
        If mid(str, ln - i + 1, 1) <> chr Then
            P(2) = ln - i + 1
            Exit For
        End If
    Next i
Else
    d = aydim(chr)
    If d = 1 And aytyp(chr) = "String" Then
        ln = Len(str)
        R = ayrge(chr)
        For i = 1 To ln
            For j = R(1, 1) To R(2, 1)
                If mid(str, i, 1) = chr(j) Then
                    b = False
                    GoTo p1
                End If
            Next j
            b = True
p1:
            If b = True Then
                P(1) = i
                Exit For
            End If
        Next i
        For i = 1 To ln
            For j = R(1, 1) To R(2, 1)
                If mid(str, ln - i + 1, 1) = chr(j) Then
                    b = False
                    GoTo p2
                End If
            Next j
            b = True
p2:
            If b = True Then
                P(2) = ln - i + 1
                Exit For
            End If
        Next i
    Else
        debug_err "Trimg", , "chr must be single value or 1-dimensional array of strings."
        Exit Function
    End If
End If
If P(2) >= P(1) And P(1) > 0 Then
    TrimG = mid(str, P(1), P(2) - P(1) + 1)
ElseIf P(2) < P(1) Then
    TrimG = ""
End If
End Function

'Remove specified characters of a string
Public Function Rechar(ByVal str As String, ByVal re As Variant) As String
Dim i As Long
If isay(re) = False And TypeName(str) = "String" Then
    Rechar = Replace(str, re, "", , , vbBinaryCompare)
ElseIf isay(re, , , 1) = True And isstr(re) = True Then
    For i = LBound(re, 1) To UBound(re, 1)
        str = Replace(str, re(i), "", , , vbBinaryCompare)
    Next i
    Rechar = str
End If
End Function

'Comma Format for numbers
Function Comma(ByVal n As Variant) As String
Dim l As Long
Dim c As Long
Dim dc As Long
Dim S As String
Dim sn As Boolean
l = Len(n)
If IsNumeric(n) = True Then
    S = CStr(n)
    dc = InStr(1, n, ".")
    sn = (InStr(1, n, "E+") <> 0)
    If sn = True Then
        Comma = n
        Exit Function
    End If
    If dc = 0 Then
        c = l
    Else
        c = dc - 1
    End If
    Do While c > 3
        S = Left(S, c - 3) & "," & Right(S, (Len(S) - c + 3))
        c = c - 3
    Loop
    Comma = S
Else
    debug_err "Comma", "nnum"
    Comma = n
    Exit Function
End If
End Function

'Month name (English version)(sf: Short form)
Public Function Monthname_E(ByVal m As Long, Optional ByVal sf As Boolean) As String
Dim Mon(1 To 12) As String
Mon(1) = "January"
Mon(2) = "February"
Mon(3) = "March"
Mon(4) = "April"
Mon(5) = "May"
Mon(6) = "June"
Mon(7) = "July"
Mon(8) = "August"
Mon(9) = "September"
Mon(10) = "October"
Mon(11) = "November"
Mon(12) = "December"
If sf = False Then
    Monthname_E = Mon(m)
Else
    Monthname_E = Left(Mon(m), 3)
End If
End Function

Public Function month_totday(ByVal m As Long, Optional ByVal yr As Long) As Long
Dim lkp(1 To 12, 1 To 2) As Long
Dim i As Long
For i = 1 To 12
    lkp(i, 1) = i
    lkp(i, 2) = 30
    If i = 1 Or i = 3 Or i = 5 Or i = 7 Or i = 8 Or i = 10 Or i = 12 Then
        lkp(i, 2) = lkp(i, 2) + 1
    ElseIf i = 2 Then
        If yr > 0 Then
            If 4 Mod i = 0 Then
                lkp(i, 2) = lkp(i, 2) - 1
            Else
                lkp(i, 2) = lkp(i, 2) - 2
            End If
        End If
    End If
Next i
month_totday = lkp(m, 2)
End Function


'Char to number (For conversion of excel column)
Public Function charnum(ByVal str As String) As Long
Dim i As Long, j As Long
Dim ln As Long
Dim n As Long
Dim char As String
Dim S(1 To 26) As String
For i = 1 To 26
    S(i) = chr(64 + i)
Next i
ln = Len(str)
str = UCase(str)
If ln > 0 Then
    For i = 1 To ln
        char = mid(str, (ln - i + 1), 1)
        For j = 1 To 26
            If char = S(j) Then
                n = n + 26 ^ (i - 1) * j
                Exit For
            End If
        Next j
    Next i
    charnum = n
Else
    debug_err "charnum", , "str is a null string, please check."
End If
End Function

'Number to String
Public Function numstr(ByVal n As Long, ByVal d As Long) As String
Dim S As String
Dim l As Long
If n > 10 ^ d Then
    debug_err numstr, , "The number is greater than" & CStr(10 ^ d) & "."
    Exit Function
End If
S = CStr(n)
l = Len(S)
Do While (l < d)
    S = "0" & S
    l = l + 1
Loop
numstr = S
End Function

'Number to Char
Public Function numchar(ByVal n As Long) As String
Dim c As String
Dim F As Long
Dim i As Long
i = 1
Do While (n > 0)
    F = (n - 1) Mod 26
    c = chr(65 + F) & c
    n = Int((n - 1) / 26)
Loop
numchar = c
End Function

'Generate all possible String List with n Characters
Public Function Genstr(ByVal str As String, ByVal n As Long) As Variant
Dim i As Long, j As Long, k As Long
Dim nstr As Long
Dim P As Long, Q As Long
Dim A() As Variant
Dim pstr As String
nstr = Len(str)
If n >= 1 Then
    P = nstr ^ n
    ReDim A(1 To P) As Variant
    For i = 1 To P
        k = i
        pstr = ""
        For j = 1 To n
            Q = ((k - 1) Mod nstr) + 1
            pstr = mid(str, Q, 1) & pstr
            k = (k - Q) \ nstr + 1
        Next j
        A(i) = pstr
    Next i
    Genstr = A
Else
    debug_err "Genstr", , "n must be greater than 0"
End If
End Function

'Generate (n) random Strings
Public Function ranstr(ByVal leng As Long, Optional n As Long, Optional ByVal ltr As Boolean, Optional ByVal num As Boolean, Optional ByVal capital As Boolean, Optional ByVal special As Boolean) As Variant
Dim i As Long, R As Long
Dim str As String
Dim strs() As Variant
Dim Ch As Long
Dim l() As String
If n = 0 Then
    n = 1
ElseIf n < 0 Then
    debug_err "ranstr", , "n must be positive integers."
    Exit Function
End If
If leng >= 1 Then
    Ch = 26
    If num = True Then Ch = Ch + 10
    If capital = True Then Ch = Ch + 26
    If special = True Then Ch = Ch + 10
    'Define list
    ReDim l(1 To Ch) As String
    Ch = 0
    If ltr = True Then
        For i = 1 To 26
            l(Ch + i) = chr(96 + i)
        Next i
        Ch = Ch + 26
    End If
    If num = True Then
        For i = 1 To 10
            l(Ch + i) = chr(47 + i)
        Next i
        Ch = Ch + 10
    End If
    If capital = True Then
        For i = 1 To 26
            l(Ch + i) = chr(64 + i)
        Next i
        Ch = Ch + 26
    End If
    If special = True Then
        l(Ch + 1) = "!"
        l(Ch + 2) = "@"
        l(Ch + 3) = "#"
        l(Ch + 4) = "$"
        l(Ch + 5) = "%"
        l(Ch + 6) = "^"
        l(Ch + 7) = "&"
        l(Ch + 8) = "*"
        l(Ch + 9) = "("
        l(Ch + 10) = ")"
        Ch = Ch + 10
    End If
    If Ch = 0 Then
        debug_err "ranstr", , "At least one of these (ltr, num, capital, special) must be chosen for generating the random strings"
        Exit Function
    End If
    'Start Generating Strings
    For i = 1 To leng
        R = Int(Rnd() * Ch + 1)
        str = str & l(R)
    Next i
    ranstr = str
Else
    debug_err "ranstr", , "Leng must be greater than or equal to 1"
End If
End Function

'Generate password list up to n characters
Public Function Genpw(ByVal n As Long, Optional ByVal num As Boolean, Optional ByVal ltr As Boolean, Optional ByVal capital As Boolean, Optional ByVal special As Boolean) As Variant
Dim i As Long, j As Long
Dim str As String
Dim ct As Long
Dim t As Long, P As Long
Dim A As Variant, F As Variant
If n >= 1 Then
    ct = 0
    str = ""
    If num = True Then
        For i = 1 To 10
            str = str & chr(47 + i)
        Next i
        ct = ct + 10
    End If
    If ltr = True Then
        For i = 1 To 26
            str = str & chr(96 + i)
        Next i
        ct = ct + 26
    End If
    If capital = True Then
        For i = 1 To 26
            str = str & chr(64 + i)
        Next i
        ct = ct + 26
    End If
    If special = True Then
        str = str & "!@#$%^&*()"
        ct = ct + 10
    End If
    If ct = 0 Then
        debug_err "Genpw", , "At least one of these (ltr, num, capital, special) must be chosen for generating the random strings"
        Exit Function
    End If
    t = 0
    For i = 1 To n
        t = t + (ct ^ i)
    Next i
    ReDim F(1 To t)
    P = 0
    For i = 1 To n
        A = Genstr(str, i)
        For j = 1 To (ct ^ i)
            F(P + j) = A(j)
        Next j
        P = P + (ct ^ i)
    Next i
    Genpw = F
Else
    debug_err "Genpw", , "n must be greater than or equal to 1."
End If
End Function

'Words count of string
Function wdct(ByVal str As String, Optional ByVal break As String, Optional unique As Boolean) As Long
Dim i As Long, k As Long
Dim ct As Long
Dim ln As Long
Dim A As Variant
If break = "" Then
    break = " "
    str = TrimG(str, " ")
End If
If unique = False Then
    i = 1
    Do Until InStr(i, str, break) = 0
        ct = ct + 1
        i = InStr(i, str, break) + 1
    Loop
    wdct = ct + 1
ElseIf unique = True Then
    A = ayuni(wDay(str, break))
    wdct = UBound(A, 1) - LBound(A, 1) + 1
End If
End Function

'Words to array (Split a word into array by break)
Public Function wDay(ByVal str As String, Optional ByVal break As String) As Variant
Dim i As Long, k As Long: k = 1
Dim ln As Long, ln2 As Long
Dim ct As Long, st As Long: st = 1
Dim S() As String
If break = "" Then break = " "
ln = Len(str)
ln2 = Len(break)
ct = wdct(str, break)
ReDim S(1 To ct)
st = 1
For i = 1 To ln
    If mid(str, i, ln2) = break Then
        If i - st > 0 Then
            S(k) = mid(str, st, i - st)
        End If
        st = i + 1
        k = k + 1
    ElseIf i = ln Then
        S(k) = mid(str, st, i - st + 1)
    End If
Next i
wDay = S
End Function

'Word Frequency table
Public Function wdfreq(ByVal str As String, Optional ByVal break As String) As Variant
Dim A As Variant
A = wDay(str, break)
wdfreq = ayFreq(A)
End Function

'Date Functions
Public Function Datefunc(ByVal Date1, Optional yr As Long, Optional mth As Long, Optional dy As Long, Optional hr As Long, Optional mn As Long, Optional sd As Long) As Date
yr = year(Date1)
mth = month(Date1)
dy = day(Date1)
hr = Hour(Date1)
mn = Minute(Date1)
sd = Second(Date1)
Datefunc = Date1
End Function

'Date Conversion
Public Function Datecon(ByVal Dat As String, ByVal BF As String, ByVal AF As String) As String
Dim i As Long
Dim m(1 To 12) As String
Dim d(1 To 3) As Long
Dim F As String
Dim fs(1 To 3) As String '1 : Day, 2: Month, 3: Year
Dim b As String
m(1) = "January"
m(2) = "February"
m(3) = "March"
m(4) = "April"
m(5) = "May"
m(6) = "June"
m(7) = "July"
m(8) = "August"
m(9) = "September"
m(10) = "October"
m(11) = "November"
m(12) = "December"
'Convert standard date format
'Month Conversion
If InStr(1, BF, "MMMM") <> 0 Then
    For i = 1 To 12
        If InStr(1, Dat, m(i)) <> 0 Then
            d(2) = i
            GoTo fd
        End If
    Next i
    debug_err "Datecon", , "Current Format specified wrongly."
ElseIf InStr(1, BF, "MMM") <> 0 Then
    For i = 1 To 12
        If InStr(1, Dat, Left(m(i), 3)) <> 0 Then
            d(2) = i
            GoTo fd
        End If
    Next i
    debug_err "Datecon", , "Current Format specified wrongly."
ElseIf InStr(1, BF, "MM") <> 0 Then
    d(2) = CLng(mid(Dat, InStr(1, BF, "MM"), 2))
ElseIf InStr(1, BF, "M") <> 0 Then
    d(2) = CLng(mid(Dat, InStr(1, BF, "M"), 1))
End If
fd:
'Day Conversion
If InStr(1, BF, "DD") <> 0 Then
    If InStr(1, BF, "MMMM") = 0 Then
        d(1) = CLng(mid(Dat, InStr(1, BF, "DD"), 2))
    ElseIf InStr(1, BF, "MMMM") <> 0 Then
    End If
ElseIf InStr(1, BF, "D") <> 0 Then
    If InStr(1, BF, "MMMM") = 0 Then
        d(1) = CLng(mid(Dat, InStr(1, BF, "D"), 1))
    End If
Else
    debug_err "Datecon", , "Current Format specified wrongly."
    Exit Function
End If
'Year Conversion
If InStr(1, BF, "YYYY") <> 0 Then
    If InStr(1, BF, "MMMM") = 0 Then
        d(3) = CLng(mid(Dat, InStr(1, BF, "YYYY"), 4))
    End If
ElseIf InStr(1, BF, "YY") <> 0 Then
    If InStr(1, BF, "MMMM") = 0 Then
        d(3) = CLng(mid(Dat, InStr(1, BF, "YY"), 2))
    End If
Else
    debug_err "Datecon", , "Current Format specified wrongly."
    Exit Function
End If
'Check date error
If d(1) < 1 Or d(1) > 31 Or d(2) < 1 Or d(2) > 12 Or d(3) < 1 Or d(3) > 2999 Then
    debug_err "Datecon", , "Wrong date."
    Exit Function
End If
'Day Conversion
If InStr(1, AF, "DD") <> 0 Then
    If d(1) < 10 Then
        fs(1) = "0" & CStr(d(1))
    ElseIf d(i) > 10 Then
        fs(1) = CStr(d(1))
    End If
ElseIf InStr(1, AF, "D") Then
    fs(1) = CStr(d(1))
End If
'Month Conversion
If InStr(1, AF, "MMMM") <> 0 Then
    fs(2) = m(d(2))
ElseIf InStr(1, AF, "MMM") <> 0 Then
    fs(2) = Left(m(d(2)), 3)
ElseIf InStr(1, AF, "MM") Then
    If d(2) < 10 Then
        fs(2) = "0" & CStr(d(2))
    ElseIf d(i) > 10 Then
        fs(2) = CStr(d(2))
    End If
ElseIf InStr(1, AF, "M") Then
    fs(2) = CStr(d(2))
End If
'Year Conversion
If InStr(1, AF, "YYYY") <> 0 Then
    If Len(d(3)) = 1 Then
        fs(3) = "200" & CStr(d(3))
    ElseIf Len(d(3)) = 2 Then
        fs(3) = "20" & CStr(d(3))
    ElseIf Len(d(3)) = 3 Then
        fs(3) = "1" & CStr(d(3))
    ElseIf Len(d(3)) = 4 Then
        fs(3) = CStr(d(3))
    End If
ElseIf InStr(1, AF, "YY") <> 0 Then
    If Len(d(3)) = 1 Then
        fs(3) = "0" & CStr(d(3))
    ElseIf Len(d(3)) = 2 Then
        fs(3) = CStr(d(3))
    ElseIf Len(d(3)) = 3 Then
        fs(3) = Left(CStr(d(3)), 2)
    End If
End If
If InStr(1, AF, "-") <> 0 Then
    b = "-"
ElseIf InStr(1, AF, "/") <> 0 Then
    b = "/"
End If
If InStr(1, AF, "D") < InStr(1, AF, "M") And InStr(1, AF, "M") < InStr(1, AF, "Y") Then
    F = fs(1) & b & fs(2) & b & fs(3)
ElseIf InStr(1, AF, "M") < InStr(1, AF, "D") And InStr(1, AF, "D") < InStr(1, AF, "Y") Then
    F = fs(2) & b & fs(1) & b & fs(3)
ElseIf InStr(1, AF, "Y") < InStr(1, AF, "M") And InStr(1, AF, "M") < InStr(1, AF, "D") Then
    F = fs(3) & b & fs(2) & b & fs(1)
Else
    debug_err "Datecon", , "Format to be converted specified wrongly."
    Exit Function
End If
Datecon = F
End Function

'XOR encryption and decryption
Function XORencryption(str As String, Pw As String) As String
Dim dstr As String
Dim SL As Long
Dim pl As Long
Dim S() As Integer
Dim P() As Integer
SL = Len(str)
pl = Len(Pw)
ReDim S(1 To SL) As Integer
ReDim P(0 To pl - 1) As Integer
For i = 1 To pl
    P(i - 1) = Asc(mid(Pw, i, 1))
Next i
For i = 1 To SL
    S(i) = Asc(mid(str, i, 1))
    S(i) = S(i) Xor P(i Mod pl)
    dstr = dstr & chr(S(i))
Next i
XORencryption = dstr
End Function

'Google Translate string (ie version)
Function str_trans_ie(ByVal str As String, ByVal langin As String, ByVal langout As String, Optional ByRef ie As Object) As String
Dim i As Long, j As Long
Dim b As Boolean
Dim CLEAN_DATA As Variant
Dim str_r As String
If objex(ie) = False Then
    Set ie = CreateObject("InternetExplorer.application")
    b = False
Else
    b = True
End If
ie.Visible = False
    ie.navigate "http://translate.google.com/#" & langin & "/" & langout & "/" & str
    Do Until ie.readyState = 4
        DoEvents
    Loop
    CLEAN_DATA = Split(Application.WorksheetFunction.Substitute(ie.Document.getElementById("result_box").innerHTML, "</SPAN>", ""), "<")
    For i = LBound(CLEAN_DATA) To UBound(CLEAN_DATA)
        str_r = str_r & Right(CLEAN_DATA(i), Len(CLEAN_DATA(i)) - InStr(CLEAN_DATA(i), ">"))
    Next
    If b = False Then
        ie.Quit
    End If
    str_trans_ie = str_r
End Function

'Google Translate string (xml version)
Function str_trans_xml(ByVal str As String, ByVal langin As String, ByVal langout As String, Optional ByRef WinHttpReq As Object) As String
Dim b As Boolean
Dim j As Long
Dim Url As String
Dim strData As String
Dim id As String
Dim CLEAN_DATA As Variant
Dim Result_data As Variant
    On Error Resume Next
    j = WinHttpReq.ResponseStatus
    If ERR.Number > 0 Then
        Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
        b = False
    Else
        b = True
    End If
    On Error GoTo 0

    'Url = "https://translate.google.com/m?hl=" & langout & "&sl=" & langin & "&tl=" & langout & "&ie=UTF-8&prev=_m&q=" & str & ""
    Url = "https://translate.google.com/m?hl=" & langin & "&sl=" & langin & "&tl=" & langout & "&prev=_m&q=" & str & ""
    
    WinHttpReq.Open "GET", Url, False
    WinHttpReq.Send
    If WinHttpReq.Status = 200 Then
        strData = WinHttpReq.responseText
        id = "<div dir=""ltr"" class=""t0"">"
        strData = mid(strData, InStr(1, strData, id, vbBinaryCompare) + Len(id))
        strData = mid(strData, 1, InStr(1, strData, "</div>", vbBinaryCompare) - 1)
        CLEAN_DATA = Split(strData, "<")
        
        For j = LBound(CLEAN_DATA) To UBound(CLEAN_DATA)
            Result_data = Result_data & Right(CLEAN_DATA(j), Len(CLEAN_DATA(j)) - InStr(CLEAN_DATA(j), ">"))
        Next
    Else
        debug_err "str_trans_xml", , "Error for the response"
    End If
    If b = False Then
        Set WinHttpReq = Nothing
    End If
    str_trans_xml = Result_data
    
'»y¨¥ªþµù:
'auto,°»´ú»y¨¥
'zh-TW,¤¤¤å(ÁcÅé)
'es,¦è¯Z¤ú¤å
'en,­^¤å
'
'tr,¤g¦Õ¨ä¤å      'af,¥¬º¸¤å
'zh-TW,¤¤¤å(ÁcÅé) 'fy,¥±§QµM¤å
'zh-CN,¤¤¤å(Â²Åé) 'be,¥Õ«XÃ¹´µ¤å
'da,¤¦³Á¤å        'lt,¥ß³³©{¤å
'eu,¤Ú´µ§J¤å      'ig,¥ì³Õ¤å
'ja,¤é¤å          'is,¦B®q¤å
'mi,¤ò§Q¤å        'hu,¦I¤ú§Q¤å
'jw,¤ö«z¤å        'id,¦L¥§¤å
'gl,¥[¨½¦è¨È¤å    'su,¦L¥§´S¥L¤å
'ca,¥[®õÃ¹¥§¨È¤å  'hi,¦L«×¤å
'kn,¥d¯Ç¹F¤å      'gu,¦L«×¥j«¢©Ô¦a¤å
'ne,¥§ªyº¸¤å      'ky,¦Nº¸¦N´µ¤å
'es,¦è¯Z¤ú¤å      'bs,ªi¦è¥§¨È
'hr,§JÃ¹®J¦è¨È¤å  'fa,ªi´µ¤å
'iw,§Æ§B¨Ó¤å      'pl,ªiÄõ¤å
'el,§ÆÃ¾¤å        'fi,ªâÄõ¤å
'hy,¨È¬ü¥§¨È¤å    'am,ªü©i«¢©Ô¤å
'az,¨È¶ë«ôµM¤å    'ar,ªü©Ô§B¤å
'ny,©_¤Á¥Ë¤å      'sq,ªüº¸¤Ú¥§¨È¤å
'bn,©s¥[©Ô¤å      'ru,«X¤å
'ps,©¬¬I¹Ï¤å      'bg,«O¥[§Q¨È¤å
'la,©Ô¤B¤å        'sd,«H¼w¤å
'lv,©Ô²æºû¨È¤å    'xh,«n«D¬_ÂÄ¤å
'fr,ªk¤å          'zu,«n«D¯ª¾|¤å
'kk,«¢ÂÄ§J¤å 'ht,®ü¦a§J¨½¶ø¤å
'cy,«Âº¸´µ¤å 'uk,¯Q§JÄõ¤å
'co,¬ì¦è¹Å¤å 'uz,¯Q¯÷§O§J¤å
'hmn,­]¤å    'ur,¯Qº¸³£¤å
'en,­^¤å     'so,¯Á°¨¨½¤å
'haw,®L«Â¦i¤å'mt,°¨¦Õ¥L¤å
'ku,®w¼w¤å   'ms,°¨¨Ó¤å
'no,®¿«Â¤å   'mk,°¨¨ä¹y¤å
'pa,®Ç¾B´¶¤å 'mg,°¨©Ô¥[´µ¤å
'th,®õ¤å     'mr,°¨©Ô¦a¤å
'ta,®õ¦Ìº¸¤å 'ml,°¨©Ô¶®©Ô©i¤å
'te,®õ¿c©T¤å 'km,°ª´Ö¤å
'eo,°ê»Ú»y¤å     'sr,¶ëº¸ºû¨È¤å
'ceb,±JÃú¤å      'yi,·N²Äºü¤å
'cs,±¶§J¤å       'et,·R¨F¥§¨È¤å
'sn,²Ð¯Ç¤å       'ga,·Rº¸Äõ¤å
'nl,²üÄõ¤å       'sv,·ç¨å¤å
'ka,³ìªv¨È¤å     'st,·æ¯Á¦«¤å
'sw,´µ¥Ë§Æ¨½¤å   'it,¸q¤j§Q¤å
'sk,´µ¬¥¥ï§J¤å   'pt,¸²µå¤ú¤å
'sl,´µ¬¥ºû¥§¨È¤å 'mn,»X¥j¤å
'tl,µá«ß»«¤å     'ha,»¨¨F¤å
'vi,¶V«n¤å       'lo,¼d¤å
'tg,¶ð¦N§J¤å     'de,¼w¤å
'my,½q¨l¤å
'lb,¿c´Ë³ù¤å
'si,¿üÄõ¤å
'yo,Àu¾|¤Ú¤å
'ko,Áú¤å
'sm,ÂÄ¼¯¨È¤å
'ro,Ã¹°¨¥§¨È¤å
'gd,Ä¬®æÄõªº»\º¸¤å
End Function

'******************************************************************************************************************************************************

'C!: Array Functions
'is an array? (option: jagged array, numeric array)
Public Function isay(ByVal ay As Variant, Optional ByVal jag As Boolean, Optional ByVal num As Boolean, Optional ByVal di As Integer, Optional ByVal iint As Boolean, Optional ByVal sign As String) As Boolean
Dim A As Variant, AA As Variant
If IsArray(ay) = True Then
    If num = True Then
        If aynum(ay) = False Then
            isay = False
            Exit Function
        End If
    End If
    If di >= 1 Then
        If di <> aydim(ay) Then
            isay = False
            Exit Function
        End If
    End If
    If jag = False Then
        If iint = False Then
            isay = True
        ElseIf iint = True Then
            For Each A In ay
                If isint(A, sign) = False Then
                    isay = False
                    Exit Function
                End If
            Next A
            isay = True
        End If
    Else
        For Each A In ay
            If IsArray(A) = True Then
                If iint = False Then
                    isay = True
                ElseIf iint = True Then
                    For Each AA In A
                        If isint(AA, sign) = False Then
                            isay = False
                            Exit Function
                        End If
                    Next AA
                End If
            End If
        Next A
    End If
Else
    isay = False
End If
End Function

'Isnumeric of an array (useful for non-array and jagged array)
Public Function aynum(ByVal ay As Variant) As Boolean
Dim A As Variant
If IsArray(ay) = False Then
    If IsNumeric(ay) = False Then
        aynum = False
        Exit Function
    End If
Else
    For Each A In ay
        If aynum(A) = False Then
            aynum = False
            Exit Function
        End If
    Next A
End If
aynum = True
End Function

'Dimension of an array
Public Function aydim(ByVal ay As Variant) As Integer
Dim i As Byte: i = 1
Dim t As Long
Do
    On Error GoTo R
    t = LBound(ay, i)
    i = i + 1
Loop
R:
aydim = i - 1
End Function

'Range of an array (Returns a 2-dimension array: Lbound vs Ubound, n th dimension)
Public Function ayrge(ByVal ay As Variant) As Variant
Dim i As Long
Dim d As Integer
Dim R As Variant
If isay(ay) = True Then
    d = aydim(ay)
    ReDim R(1 To 2, 1 To d) 'Lbound vs Ubound, nth dimension
    For i = 1 To d
        R(1, i) = LBound(ay, i)
        R(2, i) = UBound(ay, i)
    Next i
    ayrge = R
Else
    debug_err "ayrge", "nay"
End If
End Function

'Type of an array
Public Function aytyp(ByVal ay As Variant) As String
Dim i As Long, j As Long, k As Long
Dim A As Variant
Dim Typ(1 To 21) As String
Dim typc(0 To 21, 0 To 1) As String
Dim ct As Long
Dim typb(1 To 6) As Boolean '1: Empty, 2: Boolean, 3: Numeric, 4: String, 5: Date, 6: Others
typc(0, 0) = "Empty"
typc(1, 0) = "DBNull"
typc(2, 0) = "Nothing"
typc(3, 0) = "Boolean"
typc(4, 0) = "Byte"
typc(5, 0) = "Integer"
typc(6, 0) = "Short"
typc(7, 0) = "Long"
typc(8, 0) = "SByte"
typc(9, 0) = "UInteger"
typc(11, 0) = "ULong"
typc(10, 0) = "UShort"
typc(12, 0) = "Single"
typc(13, 0) = "Double"
typc(14, 0) = "Decimal"
typc(15, 0) = "Long"
typc(16, 0) = "Char"
typc(17, 0) = "String"
typc(18, 0) = "Date"
typc(19, 0) = "Object"
typc(20, 0) = "objectclass"
typc(21, 0) = "Others"
For j = 0 To 21
    typc(j, 1) = False
Next j
If IsArray(ay) = True Then
    i = 1
    For Each A In ay
        k = 1
        Do Until Typ(k) = ""
            If TypeName(A) = Typ(k) Then
                GoTo SKip
            End If
            k = k + 1
        Loop
        Typ(i) = TypeName(A)
        i = i + 1
SKip:
    Next A
    k = 1
    Do Until Typ(k) = ""
        For j = 0 To 21
            If Typ(k) = typc(j, 0) Then
                typc(j, 1) = True
                Exit For
            ElseIf j = 21 Then
                typc(j, 1) = True
            End If
        Next j
        k = k + 1
    Loop
    ct = k - 1
    If ct = 1 Then
        aytyp = Typ(ct)
    ElseIf ct > 1 Then
        For j = 0 To 21
            If typc(j, 1) = True And j >= 0 And j <= 2 Then
                typb(1) = True
            ElseIf typc(j, 1) = True And j = 3 Then
                typb(2) = True
            ElseIf typc(j, 1) = True And j >= 4 And j <= 15 Then
                typb(3) = True
            ElseIf typc(j, 1) = True And j >= 16 And j <= 17 Then
                typb(4) = True
            ElseIf typc(j, 1) = True And j = 18 Then
                typb(5) = True
            ElseIf typc(j, 1) = True And j >= 19 And j <= 21 Then
                typb(6) = True
            End If
        Next j
        If typb(6) = True Then
            If (typb(2) Or typb(3) Or typb(4) Or typb(5)) = True Then
                aytyp = "Variant"
            Else
                aytyp = "Others"
            End If
        ElseIf typb(6) = False Then
            If typb(5) = True Then
                If (typb(2) Or typb(3) Or typb(4)) = True Then
                    aytyp = "Variant"
                Else
                    aytyp = "Date"
                End If
            ElseIf typb(5) = False Then
                If typb(4) = True Then
                    aytyp = "String"
                ElseIf typb(4) = False Then
                    If typb(2) = True And typb(3) = True Then
                        aytyp = "String"
                    ElseIf typb(2) = True And typb(3) = False Then
                        aytyp = "Boolean"
                    ElseIf typb(3) = True And typb(2) = False Then
                        aytyp = "Numeric"
                    ElseIf typb(2) = False And typb(3) = False Then
                        aytyp = "Empty"
                    End If
                End If
            End If
        End If
    End If
ElseIf IsArray(ay) = False Then
    For j = 0 To 21
        If TypeName(ay) = typc(j, 0) And j >= 0 And j <= 2 Then
            aytyp = "Empty"
            GoTo Skipp
        ElseIf TypeName(ay) = typc(j, 0) And j = 3 Then
            aytyp = "Boolean"
            GoTo Skipp
        ElseIf TypeName(ay) = typc(j, 0) And j >= 4 And j <= 15 Then
            aytyp = "Numeric"
            GoTo Skipp
        ElseIf TypeName(ay) = typc(j, 0) And j >= 16 And j <= 17 Then
            aytyp = "String"
            GoTo Skipp
        ElseIf TypeName(ay) = typc(j, 0) And j = 18 Then
            aytyp = "Date"
            GoTo Skipp
        End If
    Next j
    aytyp = "Others"
Skipp:
End If
End Function

'Total count of an array
Public Function Aytct(ByVal ay As Variant, Optional jag As Boolean, Optional ety As Boolean) As Long
Dim A As Variant
Dim ct As Long
If isay(ay, True) = True And jag = True Then
    For Each A In ay
        ct = ct + Aytct(A, True, ety)
    Next A
ElseIf isay(ay) = True Then
    For Each A In ay
        If ety = False Then
            ct = ct + 1
        Else
            If IsNull(A) = False And IsEmpty(A) = False Then
                ct = ct + 1
            End If
        End If
    Next A
ElseIf isay(ay) = False Then
    If ety = False Then
        ct = 1
    Else
        If IsNull(ay) = False And IsEmpty(ay) = False Then
            ct = 1
        End If
    End If
End If
Aytct = ct
End Function

'Turn all values to 1 dimension array from an array (jagged array)
Public Function aytr1(ByVal ay As Variant) As Variant
Dim i As Long, j As Long
Dim v As Variant
Dim A() As Variant, AA As Variant
If isay(ay, True) = True Then
    ReDim A(1 To Aytct(ay, True)) As Variant
    i = 1
    For Each v In ay
        AA = aytr1(v)
        For j = 1 To Aytct(ay)
            A(i + j - 1) = AA(j)
        Next j
        i = i + Aytct(v)
    Next v
    aytr1 = A
ElseIf isay(ay) = True Then
    ReDim A(1 To Aytct(ay)) As Variant
    i = 1
    For Each v In ay
        A(i) = v
        i = i + 1
    Next v
    aytr1 = A
Else
    aytr1 = ay
End If
End Function

'Extract (1 to n - 1) dimension (up to 4 dimensions) from a n - dimensional array (up to 10 dimensions) (c: co-ordinates input to be extracted)
Function ayredim(ByVal ay As Variant, ByRef c As Variant, ByVal x As Variant, Optional ByVal st As Variant, Optional ByVal ed As Variant) As Variant
Dim i As Long
Dim d As Integer, dx As Integer
Dim R As Variant, RC As Variant, RX As Variant, rx1 As Variant
Dim A As Variant, F As Variant
d = aydim(ay)
R = ayrge(ay)
If isay(c, , True, 1) = False Then
    debug_err "ayredim", , "coord is not a numeric 1 - dimensional array."
    Exit Function
End If
RC = ayrge(c)
If d = ayct(c) Then
    For i = 1 To d
        If c(RC(1, 1) + i - 1) = "" Then c(RC(1, 1) + i - 1) = R(1, i)
    Next i
ElseIf d <> ayct(c) Then
    debug_err "ayredim", , "The count of coord does not match with the dimension of ay."
End If
For i = 1 To d
    If c(RC(1, 1) + i - 1) < R(1, i) Or c(RC(1, 1) + i - 1) > R(2, i) Then
        debug_err "Ayredim", , "At least one of the coords is out of the range of the array, please check!"
        Exit Function
    End If
Next i
If isay(x) = False Then
    dx = 1
    rx1 = x
    If x < 1 Or x > d Then
        debug_err "Ayredim", , "x specified is out of the range of the dimension of the array."
        Exit Function
    End If
    If IsMissing(st) = True Then
        st = R(1, x)
    ElseIf st < R(1, x) Or st > R(2, x) Then
        debug_err "Ayredim", , "st is out of the range of " & x & " - dimensional array."
    End If
    If IsMissing(ed) = True Then
        ed = R(2, x)
    ElseIf ed < R(1, x) Or ed > R(2, x) Then
        debug_err "Ayredim", , "ed is out of the range of " & x & " - dimensional array."
    End If
ElseIf isay(x, , True, 1) = True Then
    dx = ayct(x)
    RX = ayrge(x)
    rx1 = RX(1, 1)
    For i = RX(1, 1) To RX(2, 1)
        If x(i) < 1 Or x(i) > d Then
            debug_err "Ayredim", , "x specified is out of the range of the dimension of the array."
            Exit Function
        End If
    Next i
    If isay(st) = False Then
        If IsMissing(st) = True Then
            ReDim A(RX(1, 1) To RX(2, 1))
            For i = RX(1, 1) To RX(2, 1)
                A(i) = R(1, x(i))
            Next i
            st = A
        Else
            debug_err "Ayredim", , "st must be a 1-dimensional array."
        End If
    ElseIf isay(st, , True, 1) = False Then
        If ayct(st) = dx Then
            For i = RX(1, 1) To RX(2, 1)
                If IsMissing(st(i)) = True Then
                    st(i) = R(1, x(i))
                End If
            Next i
        Else
            debug_err "Ayredim", , "The dimension of st, ed must match with x."
        End If
    End If
    If isay(ed) = False Then
        If IsMissing(ed) = True Then
            ReDim A(RX(1, 1) To RX(2, 1))
            For i = RX(1, 1) To RX(2, 1)
                A(i) = R(2, x(i))
            Next i
            ed = A
        Else
            debug_err "Ayredim", , "ed must be a 1-dimensional array."
        End If
    ElseIf isay(ed, , True, 1) = False Then
        If ayct(ed) = dx Then
            For i = RX(1, 1) To RX(2, 1)
                If IsMissing(ed(i)) = True Then
                    ed(i) = R(2, x)
                End If
            Next i
        Else
            debug_err "Ayredim", , "The dimension of st, ed must match with x."
        End If
    End If
Else
    debug_err "ayredim", , "x must be a single value or a 1-dimensional array."
End If
F = ayredim_r(st, ed)
ayredim = ayredim_d(ay, c, x, d, RC(1, 1), dx, st, ed, rx1)
End Function
    Function ayredim_d(ByVal ay As Variant, ByVal c As Variant, ByVal x As Variant, ByVal d As Integer, ByVal S As Integer, ByVal dx As Integer, ByVal st As Variant, ByVal ed As Variant, ByVal xs As Integer) As Variant
    Dim F As Variant
    Dim i As Long, j As Long, k As Long, G As Long
    Dim xf As Long
    Dim edf As Long, stf As Long
    Dim lo As Long
    lo = LBound(c)
    If dx = 1 Then
        ReDim F(1 To ed - st + 1) As Variant
        For i = st To ed
            c(lo + x - 1) = i
            F(i - st + 1) = ayredim_c(ay, c, d, S)
        Next i
    ElseIf dx = 2 Then
        ReDim F(1 To ed(xs) - st(xs) + 1, 1 To ed(xs + 1) - st(xs + 1) + 1) As Variant
        For i = st(xs) To ed(xs)
            c(lo + x(xs) - 1) = i
            For j = st(xs + 1) To ed(xs + 1)
                c(lo + x(xs + 1) - 1) = j
                F(i - st(xs) + 1, j - st(xs + 1) + 1) = ayredim_c(ay, c, d, S)
            Next j
        Next i
    ElseIf dx = 3 Then
        ReDim F(1 To ed(xs) - st(xs) + 1, 1 To ed(xs + 1) - st(xs + 1) + 1, 1 To ed(xs + 2) - st(xs + 2) + 1) As Variant
        For i = st(xs) To ed(xs)
            c(lo + x(xs) - 1) = i
            For j = st(xs + 1) To ed(xs + 1)
                c(lo + x(xs + 1) - 1) = j
                For k = st(xs + 2) To ed(xs + 2)
                    c(lo + x(xs + 2) - 1) = k
                    F(i - st(xs) + 1, j - st(xs + 1) + 1, k - st(xs + 2) + 1) = ayredim_c(ay, c, d, S)
                Next k
            Next j
        Next i
    ElseIf dx = 4 Then
        ReDim F(1 To ed(xs) - st(xs) + 1, 1 To ed(xs + 1) - st(xs + 1) + 1, 1 To ed(xs + 2) - st(xs + 2) + 1, 1 To ed(xs + 3) - st(xs + 3) + 1) As Variant
        For i = st(xs) To ed(xs)
            c(lo + x(xs) - 1) = i
            For j = st(xs + 1) To ed(xs + 1)
                c(lo + x(xs + 1) - 1) = j
                For k = st(xs + 2) To ed(xs + 2)
                    c(lo + x(xs + 2) - 1) = k
                    For G = st(xs + 3) To ed(xs + 3)
                        c(lo + x(xs + 3) - 1) = G
                        F(i - st(xs) + 1, j - st(xs + 1) + 1, k - st(xs + 2) + 1, G - st(xs + 3) + 1) = ayredim_c(ay, c, d, S)
                    Next G
                Next k
            Next j
        Next i
    End If
    ayredim_d = F
    End Function
    Function ayredim_c(ByVal ay As Variant, ByVal c As Variant, ByVal d As Integer, ByVal S As Integer) As Variant
    If d = 1 Then
        ayredim_c = ay(c(S))
    ElseIf d = 2 Then
        ayredim_c = ay(c(S), c(S + 1))
    ElseIf d = 3 Then
        ayredim_c = ay(c(S), c(S + 1), c(S + 2))
    ElseIf d = 4 Then
        ayredim_c = ay(c(S), c(S + 1), c(S + 2), c(S + 3))
    ElseIf d = 5 Then
        ayredim_c = ay(c(S), c(S + 1), c(S + 2), c(S + 3), c(S + 4))
    ElseIf d = 6 Then
        ayredim_c = ay(c(S), c(S + 1), c(S + 2), c(S + 3), c(S + 4), c(S + 5))
    ElseIf d = 7 Then
        ayredim_c = ay(c(S), c(S + 1), c(S + 2), c(S + 3), c(S + 4), c(S + 5), c(S + 6))
    ElseIf d = 8 Then
        ayredim_c = ay(c(S), c(S + 1), c(S + 2), c(S + 3), c(S + 4), c(S + 5), c(S + 6), c(S + 7))
    ElseIf d = 9 Then
        ayredim_c = ay(c(S), c(S + 1), c(S + 2), c(S + 3), c(S + 4), c(S + 5), c(S + 6), c(S + 7), c(S + 8))
    ElseIf d = 10 Then
        ayredim_c = ay(c(S), c(S + 1), c(S + 2), c(S + 3), c(S + 4), c(S + 5), c(S + 6), c(S + 7), c(S + 8), c(S + 9))
    End If
    End Function
    Function ayredim_r(ByVal st As Variant, ByVal ed As Variant) As Variant
    Dim A As Variant
    Dim R As Variant
    If isay(st) = True Then
        R = ayrge(st)
    End If
    If isay(st) = False Then
        ReDim A(1 To (ed - st + 1)) As Variant
    ElseIf isay(st, , True, 2) = True Then
        ReDim A(1 To ed(R(1, 1)) - st(R(1, 1)) + 1, 1 To ed(R(1, 1) + 1) - st(R(1, 1) + 1) + 1) As Variant
    ElseIf isay(st, , True, 3) = True Then
        ReDim A(1 To ed(R(1, 1)) - st(R(1, 1)) + 1, 1 To ed(R(1, 1) + 1) - st(R(1, 1) + 1) + 1, 1 To ed(R(1, 1) + 2) - st(R(1, 1) + 2) + 1) As Variant
    ElseIf isay(st, , True, 4) = True Then
        ReDim A(1 To ed(R(1, 1)) - st(R(1, 1)) + 1, 1 To ed(R(1, 1) + 1) - st(R(1, 1) + 1) + 1, 1 To ed(R(1, 1) + 2) - st(R(1, 1) + 2) + 1, 1 To ed(R(1, 1) + 3) - st(R(1, 1) + 3) + 1) As Variant
    ElseIf isay(st, , True, 5) = True Then
        ReDim A(1 To ed(R(1, 1)) - st(R(1, 1)) + 1, 1 To ed(R(1, 1) + 1) - st(R(1, 1) + 1) + 1, 1 To ed(R(1, 1) + 3) - st(R(1, 1) + 3) + 1, 1 To ed(R(1, 1) + 3) - st(R(1, 1) + 3) + 1, 1 To ed(R(1, 1) + 3) - st(R(1, 1) + 3) + 1) As Variant
    End If
    ayredim_r = A
    End Function
    
'Search position by text of an array (Before: aycol) (up to 6-dimensional array)
Public Function aypos(ByVal ay As Variant, ByVal text As String, Optional ByVal di As Integer, Optional ByVal pos As Variant) As Long
Dim i As Long
Dim d As Integer
Dim R As Variant, RP As Variant
If isay(ay) = True Then
    d = aydim(ay)
    R = ayrge(ay)
    If d > 1 Then
        If isay(pos, , , 1) = False Then
            debug_err "aypos", "1d - pos"
            Exit Function
        Else
            If (UBound(pos, 1) - LBound(pos, 1) + 1) <> d Then
                debug_err "aypos", , "pos is wrongly specified. please check."
                Exit Function
            End If
        End If
        If di < 1 Or di > d Then
            debug_err "aypos", , "di is out of the range."
            Exit Function
        End If
        RP = ayrge(pos)
    End If
    If d = 1 Then
        For i = R(1, 1) To R(2, 1)
            If Trim(CStr(ay(i))) = text Then
                aypos = i
            End If
        Next i
    ElseIf d > 1 Then
        For i = R(1, d) To R(2, d)
            pos(RP(1, 1) + di - 1) = i
            If aypos_s(ay, d, pos, RP) = text Then
                aypos = i
            End If
        Next i
    Else
        aypos = -1
        debug_err "aypos", , text & " not found in this array."
    End If
Else
    debug_err "aypos", "NAY"
End If
End Function
    Public Function aypos_s(ByVal ay As Variant, ByVal d As Integer, ByVal pos As Variant, ByVal RP As Variant) As String
    If d = 2 Then
        aypos_s = Trim(CStr(ay(pos(RP(1, 1)), pos(RP(1, 1) + 1))))
    ElseIf d = 3 Then
        aypos_s = Trim(CStr(ay(pos(RP(1, 1)), pos(RP(1, 1) + 1), pos(RP(1, 1) + 2))))
    ElseIf d = 4 Then
        aypos_s = Trim(CStr(ay(pos(RP(1, 1)), pos(RP(1, 1) + 1), pos(RP(1, 1) + 2), pos(RP(1, 1) + 3))))
    ElseIf d = 5 Then
        aypos_s = Trim(CStr(ay(pos(RP(1, 1)), pos(RP(1, 1) + 1), pos(RP(1, 1) + 2), pos(RP(1, 1) + 3), pos(RP(1, 1) + 4))))
    ElseIf d = 6 Then
        aypos_s = Trim(CStr(ay(pos(RP(1, 1)), pos(RP(1, 1) + 1), pos(RP(1, 1) + 2), pos(RP(1, 1) + 3), pos(RP(1, 1) + 4), pos(RP(1, 1) + 5))))
    End If
    End Function

'Total summation of an array
Public Function Aytsum(ByVal ay As Variant) As Double
Dim v As Variant
If isay(ay, , True) = True Then
    For Each v In ay
        Aytsum = Aytsum + CDbl(v)
    Next v
ElseIf IsNumeric(ay) = True Then
    Aytsum = ay
Else
    debug_err "aytsum", "nnumay"
    Exit Function
End If
End Function

'Total Maximum of an array
Public Function aytmax(ByVal ay As Variant) As Variant
aytmax = aymax(aytr1(ay))
End Function

'Total minimum of an array
Public Function aytmin(ByVal ay As Variant) As Variant
aytmin = aymin(aytr1(ay))
End Function

'Is array equal? (up to 5 dimensional array)
Public Function ayequal(ByVal ay1 As Variant, ByVal ay2 As Variant) As Boolean
Dim i As Long
Dim d1 As Byte, d2 As Byte
Dim r1 As Variant, r2 As Variant
d1 = aydim(ay1)
d2 = aydim(ay2)
If d1 = d2 Then
    If d1 = 0 Then
        If ay1 = ay2 Then
            ayequal = True
        End If
    ElseIf d1 > 0 Then
        r1 = ayrge(ay1)
        r2 = ayrge(ay2)
        For i = 1 To d1
            If r1(1, i) <> r2(1, i) Or r1(2, i) <> r2(2, i) Then
                ayequal = False
                Exit Function
            End If
        Next i
        ayequal = ayequal_d(ay1, ay2, r1, d1, , d1)
    End If
Else
    ayequal = False
End If
End Function
    Public Function ayequal_d(ByVal ay1 As Variant, ByVal ay2 As Variant, ByVal R As Variant, ByVal d As Byte, Optional ByVal P As Variant, Optional ByVal pd As Byte) As Boolean
    Dim i As Long
    Dim u As Long
    If pd = 1 Then
        u = UBound(R, 2)
        For i = R(1, u) To R(2, u)
            If IsArray(P) = False Then
                P = i
            ElseIf IsArray(P) = True Then
                P(u) = i
            End If
            If ayequal_c(ay1, ay2, d, P) = False Then
                ayequal_d = False
                Exit Function
            End If
        Next i
    ElseIf pd > 1 Then
        If IsArray(P) = False Then
            ReDim P(1 To d) As Variant
        End If
        pd = pd - 1
        u = LBound(R, 2) + (UBound(R, 2) - pd - 1)
        For i = R(1, u) To R(2, u)
            P(u) = i
            If ayequal_d(ay1, ay2, R, d, P, pd) = False Then
                Exit Function
            End If
        Next i
    End If
    ayequal_d = True
    End Function
    'compare
    Public Function ayequal_c(ByVal ay1 As Variant, ByVal ay2 As Variant, ByVal d As Byte, ByVal P As Variant) As Boolean
    If d = 1 Then
        If ay1(P) = ay2(P) Then
            ayequal_c = True
        End If
    ElseIf d = 2 Then
        If ay1(P(1), P(2)) = ay2(P(1), P(2)) Then
            ayequal_c = True
        End If
    ElseIf d = 3 Then
        If ay1(P(1), P(2), P(3)) = ay2(P(1), P(2), P(3)) Then
            ayequal_c = True
        End If
    ElseIf d = 4 Then
        If ay1(P(1), P(2), P(3), P(4)) = ay2(P(1), P(2), P(3), P(4)) Then
            ayequal_c = True
        End If
    ElseIf d = 5 Then
        If ay1(P(1), P(2), P(3), P(4), P(5)) = ay2(P(1), P(2), P(3), P(4), P(5)) Then
            ayequal_c = True
        End If
    End If
    End Function
'Array Clear
Public Function ayclear(ay As Variant) As Variant
Dim v As Variant
For Each v In ay
    v = ""
Next v
ayclear = ay
End Function

'Array operations (up to 2 dimensions)
Public Function ayopt(ByVal ay1 As Variant, ByVal ay2 As Variant, ByVal opt As String) As Variant
Dim i As Long, j As Long
Dim ayrge1 As Variant, ayrge2 As Variant
Dim d As Long
Dim A As Variant
If aydim(ay1) = aydim(ay2) Then
    d = aydim(ay1)
    If d = 0 Then
        ayopt = ay1 + ay2
    Else
        ayrge1 = ayrge(ay1)
        ayrge2 = ayrge(ay2)
        If ayequal(ayrge1, ayrge2) = True Then
            If d = 1 Then
                ReDim A(ayrge1(1, 1) To ayrge1(2, 1)) As Variant
                For i = ayrge1(1, 1) To ayrge1(2, 1)
                    If LCase(opt) = "add" Then
                        A(i) = ay1(i) + ay2(i)
                    ElseIf LCase(opt) = "sub" Then
                        A(i) = ay1(i) - ay2(i)
                    ElseIf LCase(opt) = "multi" Then
                        A(i) = ay1(i) * ay2(i)
                    ElseIf LCase(opt) = "div" Then
                        A(i) = ay1(i) / ay2(i)
                    Else
                        debug_err "ayadd", , "Please specify the operation to be done."
                        Exit Function
                    End If
                Next i
            ElseIf d = 2 Then
                ReDim A(ayrge1(1, 1) To ayrge1(2, 1), ayrge1(1, 2) To ayrge1(2, 2)) As Variant
                For i = ayrge1(1, 1) To ayrge1(2, 1)
                    For j = ayrge1(1, 2) To ayrge1(2, 2)
                        If LCase(opt) = "add" Then
                            A(i, j) = ay1(i, j) + ay2(i, j)
                        ElseIf LCase(opt) = "sub" Then
                            A(i, j) = ay1(i, j) - ay2(i, j)
                        ElseIf LCase(opt) = "multi" Then
                            A(i, j) = ay1(i, j) * ay2(i, j)
                        ElseIf LCase(opt) = "div" Then
                            A(i, j) = ay1(i, j) / ay2(i, j)
                        Else
                            debug_err "ayadd", , "Please specify the operation to be done."
                            Exit Function
                        End If
                    Next j
                Next i
            End If
            ayopt = A
        Else
            debug_err "ayadd", , "The two arrays are not equal."
        End If
    End If
Else
    debug_err "ayadd", , "The dimensions of the arrays are not equal."
End If
End Function

'Arrays addition
Public Function ayadd(ByVal ay1 As Variant, ByVal ay2 As Variant) As Variant
ayadd = ayopt(ay1, ay2, "add")
End Function

'Arrays Substraction
Public Function aysub(ByVal ay1 As Variant, ByVal ay2 As Variant) As Variant
aysub = ayopt(ay1, ay2, "sub")
End Function

'Arrays Multiplication
Public Function aymulti(ByVal ay1 As Variant, ByVal ay2 As Variant) As Variant
ayadd = ayopt(ay1, ay2, "multi")
End Function

'Arrays Division
Public Function aydiv(ByVal ay1 As Variant, ByVal ay2 As Variant) As Variant
ayadd = ayopt(ay1, ay2, "div")
End Function
    
'List all differences of 2 arrays (for same dimensions) (If no differences returns null value)
Public Function aydiff(ByVal ay1 As Variant, ByVal ay2 As Variant) As Variant
Dim i As Long, j As Long
Dim d1 As Byte, d2 As Byte
Dim r1 As Variant, r2 As Variant
Dim G As Long
Dim c As Long
Dim l As Variant, LF As Variant
d1 = aydim(ay1)
d2 = aydim(ay2)
If d1 = d2 Then
    If d1 = 0 Then
        If ay1 <> ay2 Then
            aydiff = ay1
        ElseIf ay1 = ay2 Then
            aydiff = ""
        End If
    ElseIf d1 > 0 Then
        r1 = ayrge(ay1)
        r2 = ayrge(ay2)
        For i = 1 To d1
            If r1(1, i) <> r2(1, i) Or r1(2, i) <> r2(2, i) Then
                debug_err "aydiff", , "The dimensions of the arrays are not equal."
                Exit Function
            End If
        Next i
        c = Aytct(ay1)
        If d1 = 1 Then
            ReDim l(1 To c) As Variant
        ElseIf d1 > 1 Then
            ReDim l(1 To c, 1 To d1) As Variant
        End If
        aydiff_d ay1, ay2, r1, d1, , d1, l, G
        If G > 0 Then
            If d1 = 1 Then
                ReDim LF(1 To G) As Variant
                For i = 1 To G
                    LF(i) = l(i)
                Next i
            ElseIf d1 > 1 Then
                ReDim LF(1 To G, 1 To d1) As Variant
                For i = 1 To G
                    For j = 1 To d1
                        LF(i, j) = l(i, j)
                    Next j
                Next i
            End If
            aydiff = LF
        ElseIf G = 0 Then
            aydiff = ""
        End If
    End If
Else
    debug_err "aydiff", , "The dimensions of the arrays are not equal."
End If
End Function
    Public Function aydiff_d(ByVal ay1 As Variant, ByVal ay2 As Variant, ByVal R As Variant, ByVal d As Byte, Optional ByVal P As Variant, Optional ByVal pd As Byte, Optional l As Variant, Optional G As Long) As Boolean
    Dim i As Long
    Dim u As Long
    If pd = 1 Then
        u = UBound(R, 2)
        For i = R(1, u) To R(2, u)
            If IsArray(P) = False Then
                P = i
            ElseIf IsArray(P) = True Then
                P(u) = i
            End If
            aydiff_c ay1, ay2, d, P, l, G
        Next i
    ElseIf pd > 1 Then
        If IsArray(P) = False Then
            ReDim P(1 To d) As Variant
        End If
        pd = pd - 1
        u = LBound(R, 2) + (UBound(R, 2) - pd - 1)
        For i = R(1, u) To R(2, u)
            P(u) = i
            aydiff_d ay1, ay2, R, d, P, pd, l, G
        Next i
    End If
    aydiff_d = True
    End Function
    'compare
    Public Function aydiff_c(ByVal ay1 As Variant, ByVal ay2 As Variant, ByVal d As Byte, ByVal P As Variant, l As Variant, G As Long) As Boolean
    Dim i As Long
    If d = 1 Then
        If ay1(P) <> ay2(P) Then
            G = G + 1
            l(G) = P
        End If
    ElseIf d = 2 Then
        If ay1(P(1), P(2)) <> ay2(P(1), P(2)) Then
            G = G + 1
            For i = 1 To d
                l(G, i) = P(i)
            Next i
        End If
    ElseIf d = 3 Then
        If ay1(P(1), P(2), P(3)) <> ay2(P(1), P(2), P(3)) Then
            G = G + 1
            For i = 1 To d
                l(G, i) = P(i)
            Next i
        End If
    ElseIf d = 4 Then
        If ay1(P(1), P(2), P(3), P(4)) <> ay2(P(1), P(2), P(3), P(4)) Then
            For i = 1 To d
                l(G, i) = P(i)
            Next i
        End If
    ElseIf d = 5 Then
        If ay1(P(1), P(2), P(3), P(4), P(5)) <> ay2(P(1), P(2), P(3), P(4), P(5)) Then
            For i = 1 To d
                l(G, i) = P(i)
            Next i
        End If
    End If
    End Function
    
'Generate permutation (Heap's Algorithm - recursive format)
Public Function permut_r(ByVal n As Long, A As Variant, Optional F As Variant, Optional k As Long)
Dim b As Boolean
Dim i As Long, j As Long
Dim R As Variant
R = ayrge(A)
If IsArray(F) = False Then
    ReDim F(1 To Npr(n, n), R(1, 1) To R(2, 1))
    b = True
    k = 1
End If
If n = 1 Then
    For j = R(1, 1) To R(2, 1)
        F(k, j) = A(j)
    Next j
    k = k + 1
Else
    For i = 0 To (n - 1)
        permut_r n - 1, A, F, k
        If n Mod 2 = 0 Then
            swap A(R(1, 1)), A(R(1, 1) + n - 1)
        ElseIf n Mod 2 = 1 Then
            swap A(R(1, 1) + i), A(R(1, 1) + n - 1)
        End If
    Next i
End If
If b = True Then
    permut_r = F
End If
End Function

'Generate permutation (Heap's Algorithm - non-recursive format)
Public Function permut(ByVal A As Variant) As Variant
Dim n As Long
Dim c() As Long
Dim F As Variant
Dim R As Variant
Dim i As Long, j As Long, k As Long
R = ayrge(A)
n = R(2, 1) - R(1, 1) + 1
ReDim c(n) As Long
ReDim F(1 To Npr(n, n), R(1, 1) To R(2, 1)) As Variant
k = 1
For j = R(1, 1) To R(2, 1)
    F(k, j) = A(j)
Next j
k = k + 1
i = 0
Do While i < n
    If c(i) < i Then
        If i Mod 2 = 0 Then
            swap A(R(1, 1)), A(R(1, 1) + i)
        Else
            swap A(R(1, 1) + c(i)), A(R(1, 1) + i)
        End If
        For j = R(1, 1) To R(2, 1)
            F(k, j) = A(j)
        Next j
        k = k + 1
        c(i) = c(i) + 1
        i = 0
    Else
        c(i) = 0
        i = i + 1
    End If
Loop
permut = F
End Function

'Generate permutation (String) (Heap's Algorithm - non-recursive format)
Public Function permut_s(ByVal str As String) As Variant
Dim i As Long, j As Long
Dim l As Long
Dim S As Variant
Dim P As Variant
Dim PS As Variant
l = Len(str)
ReDim S(1 To l) As Variant
For i = 1 To l
    S(i) = mid(str, i, 1)
Next i
P = permut(S)
ReDim PS(1 To UBound(P, 1)) As Variant
For i = 1 To UBound(P, 1)
    For j = 1 To l
        PS(i) = PS(i) & P(i, j)
    Next j
Next i
permut_s = PS
End Function

    
'******************************************************************************************************************************************************

'C1!: 1-Dimension array Functions
'Count of 1 dimension array
Public Function ayct(ByVal ay As Variant, Optional ByVal unique As Boolean) As Long
Dim i As Long
Dim R As Variant
Dim ct As Double
If isay(ay, , , 1) = True Then
    If unique = False Then
        R = ayrge(ay)
        For i = R(1, 1) To R(2, 1)
            ct = ct + 1
        Next i
        ayct = ct
    Else
        R = ayuni(ay)
        ayct = UBound(R, 1) - LBound(R, 1) + 1
    End If
Else
    debug_err "ayct", "1d"
End If
End Function

'Summation of 1 dimension array
Public Function aysum(ByVal ay As Variant) As Double
Dim i As Long
Dim R As Variant
Dim sum As Double
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        R = ayrge(ay)
        For i = R(1, 1) To R(2, 1)
            sum = sum + CDbl(ay(i))
        Next i
        aysum = sum
    Else
        debug_err "aysum", "nnumay"
    End If
Else
    debug_err "aysum", "1d"
End If
End Function

'(Weighted) Arithmetic mean of 1 dimension array
Public Function ayamean(ByVal ay As Variant, Optional ByVal Wt As Variant) As Double
Dim i As Long
Dim RA As Variant, RW As Variant
Dim WA() As Variant
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        If isay(Wt) = False Then
            ayamean = aysum(ay) / ayct(ay)
        Else
            If isay(Wt, , , 1) = True Then
                If isay(Wt, , True) = True Then
                    RA = ayrge(ay)
                    RW = ayrge(ay)
                    If RA(1, 1) = RW(1, 1) And RA(2, 1) = RW(2, 1) Then
                        ReDim WA(RA(1, 1) To RA(2, 1)) As Variant
                        For i = RA(1, 1) To RA(2, 1)
                            WA(i) = ay(i) * Wt(i)
                        Next i
                        ayamean = aysum(WA) / aysum(Wt)
                    Else
                        debug_err "ayamean", , "The boundaries of ay and wt are not match, Please check!"
                    End If
                Else
                    debug_err " ayamean", "nnum - wt"
                End If
            Else
                debug_err "ayamean", , "wt is not 1-dimension array"
            End If
        End If
    Else
        debug_err "ayamean", "nnumay"
    End If
Else
    debug_err "ayamean", "1d"
End If
End Function

'Mean of 1 dimension array
Public Function aymean(ByVal ay As Variant, Optional ByVal Wt As Variant) As Double
aymean = ayamean(ay, Wt)
End Function

'Geometric mean of 1 dimension array
Public Function aygmean(ByVal ay As Variant) As Double
Dim i As Long
Dim R As Variant
Dim t As Long
Dim P As Double
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        R = ayrge(ay)
        t = ayct(ay)
        P = 1
        For i = R(1, 1) To R(2, 1)
            P = P * CDbl(ay(i))
        Next i
        aygmean = P ^ (1 / t)
    Else
        debug_err "aygmean", "nnum"
    End If
Else
    debug_err "aygmean", "1d"
End If
End Function

'Harmonic mean of 1 dimension array
Public Function ayhmean(ByVal ay As Variant) As Double
Dim i As Long
Dim R As Variant
Dim t As Long
Dim sum As Double
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        R = ayrge(ay)
        t = ayct(ay)
        For i = R(1, 1) To R(2, 1)
            sum = sum + (1 / CDbl(ay(i)))
        Next i
        ayhmean = t / sum
    Else
        debug_err "ayhmean", "nnum"
    End If
Else
    debug_err "ayhmean", "1d"
End If
End Function

'Power mean of 1 dimension array
Public Function aypmean(ByVal ay As Variant, ByVal m As Long) As Double
Dim i As Long
Dim R As Variant
Dim n As Long
Dim sum As Double
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        R = ayrge(ay)
        n = ayct(ay)
        For i = R(1, 1) To R(2, 1)
            sum = sum + CDbl(ay(i)) ^ m
        Next i
        aypmean = (sum / n) ^ (1 / m)
    Else
        debug_err "aypmean", "nnum"
    End If
Else
    debug_err "ayhmean", "1d"
End If
End Function

'Quadratic mean of 1 dimension array
Public Function ayqmean(ByVal ay As Variant) As Double
Dim i As Long
Dim R As Variant
Dim n As Long
Dim sum As Double
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        ayqmean = aypmean(ay, 2)
    Else
        debug_err "ayqmean", "nnum"
    End If
Else
    debug_err "ayqmean", "1d"
End If
End Function

'Start of sorting

'sort a 1 dimension array (for options)
Public Function aysort_options(ByVal ay As Variant, ByVal method As String, Optional ByVal desc As Boolean, Optional ByVal str As Boolean, Optional ByVal Title As Boolean) As Variant
Dim i As Long
Dim R As Variant, RA As Variant
Dim A As Variant, AT As Variant, AB As Variant, AN As Variant
R = ayrge(ay)
If str = False Then
    str = Not (aynum(ay))
End If
RA = R
If Title = False Then
    AT = ay
ElseIf Title = True Then
    RA(2, 1) = RA(2, 1) - 1
    ReDim AT(RA(1, 1) To RA(2, 1)) As Variant
    For i = RA(1, 1) To RA(2, 1)
        AT(i) = ay(i + 1)
    Next i
End If
'Core Sorting
    AT = aysort_copt(AT, method, str)
'End of Core sorting
If desc = False Then
    AB = AT
ElseIf desc = True Then
    ReDim AD(RA(1, 1) To RA(2, 1)) As Variant
    For i = RA(1, 1) To RA(2, 1)
        AD(i) = AT(RA(2, 1) - i + RA(1, 1))
    Next i
    AB = AD
End If
If Title = False Then
    aysort_options = AB
ElseIf Title = True Then
    ReDim AN(R(1, 1) To R(2, 1)) As Variant
    AN(R(1, 1)) = ay(R(1, 1))
    For i = RA(1, 1) To RA(2, 1)
        AN(i + 1) = AB(i)
    Next i
    aysort_options = AN
End If
End Function

'sort a 1 dimension array (for core sorting)
Public Function aysort_copt(ByVal AT As Variant, ByVal method As String, Optional ByVal str As Boolean) As Variant
If method = "Select" Then
    aysort_copt = aysort_select_c(AT, str)
ElseIf method = "Bubble" Then
    aysort_copt = aysort_bubble_c(AT, str)
ElseIf method = "Insert" Then
    aysort_copt = aysort_quick_c(AT, str)
ElseIf method = "Merge" Then
    aysort_copt = aysort_merge_c(AT, str)
ElseIf method = "Quick" Then
    aysort_copt = aysort_quick_c(AT, str)
ElseIf method = "Comb" Then
    aysort_copt = aysort_comb_c(AT, str)
Else
    debug_err "aysort_options", , "Method not specified, please check."
End If
End Function

'sort a 1 dimension array (Selection sort)
Public Function aysort_select(ByVal ay As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Boolean, Optional ByVal Title As Boolean) As Variant
If isay(ay, , , 1) = True Then
    aysort_select = aysort_options(ay, "Select", desc, str, Title)
Else
    debug_err "aysort_select", "1d"
End If
End Function

'Core Sorting - Selection sort
Function aysort_select_c(ByVal ay As Variant, Optional str As Boolean) As Variant
Dim b As Boolean
Dim x As Long, y As Long
Dim R As Variant
Dim t As Variant
R = ayrge(ay)
For x = R(1, 1) To R(2, 1)
    For y = x To R(2, 1)
        If str = False Then
            b = CDbl(ay(y)) < CDbl(ay(x))
        ElseIf str = True Then
            b = UCase(ay(y)) < UCase(ay(x))
        End If
        If b = True Then
            t = ay(x)
            ay(x) = ay(y)
            ay(y) = t
        End If
    Next y
Next x
aysort_select_c = ay
End Function

'sort a 1 dimension array (Bubble sort)
Public Function aysort_bubble(ByVal ay As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Boolean, Optional ByVal Title As Boolean) As Variant
If isay(ay, , , 1) = True Then
    aysort_bubble = aysort_options(ay, "Bubble", desc, str, Title)
Else
    debug_err "aysort_bubble", "1d"
End If
End Function

'Bubble sort
Public Function aysort_bubble_c(ByVal ay As Variant, Optional ByVal str As Boolean) As Variant
Dim b As Boolean, E As Boolean, ba As Boolean
Dim i As Long
Dim R As Variant
Dim t As Variant
R = ayrge(ay)
Do While E = False
    b = False
    For i = R(1, 1) To (R(2, 1) - 1)
        If str = False Then
            ba = CDbl(ay(i)) > CDbl(ay(i + 1))
        ElseIf str = True Then
            ba = CStr(ay(i)) > CStr(ay(i + 1))
        End If
        If ba = True Then
            t = ay(i)
            ay(i) = ay(i + 1)
            ay(i + 1) = t
            If b = False Then
                b = True
            End If
        End If
    Next i
    If b = True Then
        E = False
    ElseIf b = False Then
        E = True
    End If
Loop
aysort_bubble_c = ay
End Function

'sort a 1 dimension array (Selection sort)
Public Function aysort_insert(ByVal ay As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Boolean, Optional ByVal Title As Boolean) As Variant
If isay(ay, , , 1) = True Then
    aysort_insert = aysort_options(ay, "Insert", desc, str, Title)
Else
    debug_err "aysort_insert", "1d"
End If
End Function

'Core Sorting - Insertion sort
Function aysort_insert_ca(ByVal ay As Variant, Optional str As Boolean) As Variant
Dim i As Long
Dim x As Long, y As Long, u As Long
Dim R As Variant
Dim c1 As Variant, c2 As Variant, t As Variant
R = ayrge(ay)
For x = (R(1, 1) + 1) To R(2, 1)
    If str = False Then
        c1 = CDbl(ay(x))
    ElseIf str = True Then
        c1 = UCase(ay(x))
    End If
    For y = (x - 1) To R(1, 1) Step -1
        If str = False Then
            c2 = CDbl(ay(y))
        ElseIf str = True Then
            c2 = UCase(ay(y))
        End If
        If c1 < c2 Then
            If y = R(1, 1) Then
                u = y
            End If
        ElseIf c1 >= c2 Then
            u = y + 1
            Exit For
        End If
    Next y
    t = ay(x)
    For i = x To (u + 1) Step -1
        ay(i) = ay(i - 1)
    Next i
    ay(u) = t
Next x
aysort_insert_ca = ay
End Function

'Sort a 1 dimension array (Merge sort)
Public Function aysort_merge(ByVal ay As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Boolean, Optional ByVal Title As Boolean) As Variant
If isay(ay, , , 1) = True Then
    aysort_merge = aysort_options(ay, "Merge", desc, str, Title)
Else
    debug_err "aysort_merge", "1d"
End If
End Function

'Merge sort
Public Function aysort_merge_c(ByVal ay As Variant, Optional ByVal str As Boolean) As Variant
Dim ay1 As Variant
Dim ay2 As Variant
Dim R(1 To 2) As Variant
R(1) = LBound(ay, 1)
R(2) = UBound(ay, 1)
If R(2) - R(1) = 0 Then
    aysort_merge_c = ay
ElseIf R(2) - R(1) = 1 Then
    aysort_merge_c = aysort_bubble_c(ay, str)
ElseIf R(2) - R(1) > 1 Then
    aysplit ay, ay1, ay2
    ay1 = aysort_merge_c(ay1, str)
    ay2 = aysort_merge_c(ay2, str)
    aysort_merge_c = aymerge(ay1, ay2, str)
End If
End Function

    'Split an array into 2 array (by half)
    Public Function aysplit(ay As Variant, ay1 As Variant, ay2 As Variant) As Long
    Dim i As Long
    Dim R(1 To 2) As Variant
    Dim mid As Long
    R(1) = LBound(ay, 1)
    R(2) = UBound(ay, 1)
    mid = Int((R(1) + R(2)) / 2)
    ReDim ay1(R(1) To mid) As Variant
    ReDim ay2(R(1) To (R(1) + R(2) - (mid + 1))) As Variant
    For i = R(1) To R(2)
        If i <= mid Then
            ay1(i) = ay(i)
        ElseIf i > mid Then
            ay2(R(1) + i - (mid + 1)) = ay(i)
        End If
    Next i
    aysplit = (mid - R(1) + 1)
    End Function

    'Merge two array into one array with sorting
    Public Function aymerge(ay1 As Variant, ay2 As Variant, Optional ByVal str As Boolean) As Variant
    Dim b As Integer
    Dim i As Long, j As Long, x As Long, y As Long
    Dim R(1 To 2, 1 To 2) As Long
    Dim A As Variant
    Dim c1 As Variant, c2 As Variant
    On Error Resume Next
    R(1, 1) = LBound(ay1, 1)
    R(1, 2) = UBound(ay1, 1)
    R(2, 1) = LBound(ay2, 1)
    R(2, 2) = UBound(ay2, 1)
    i = R(1, 2) - R(1, 1) + 1
    j = R(2, 2) - R(2, 1) + 1
    If j >= R(2, 1) Then
        i = i + (R(2, 2) - R(2, 1) + 1)
    End If
    ReDim A(R(1, 1) To i) As Variant
    x = R(1, 1)
    y = R(2, 1)
    For j = R(1, 1) To i Step 0
        If x > (R(1, 2) - R(1, 1) + 1) And y > (R(2, 2) - R(2, 1) + 1) Then
            aymerge = A
            Exit Function
        End If
        If str = False Then
            c1 = CDbl(ay1(x))
            c2 = CDbl(ay2(y))
        ElseIf str = True Then
            c1 = CStr(ay1(x))
            c2 = CStr(ay2(y))
        End If
        If c1 = c2 And (x <= (R(1, 2) - R(1, 1) + 1) And y <= (R(2, 2) - R(2, 1) + 1)) Then
            A(j) = ay1(x)
            A(j + 1) = ay2(y)
            y = y + 1
            x = x + 1
            j = j + 2
        ElseIf (c1 < c2 Or y > (R(2, 2) - R(2, 1) + 1)) And x <= (R(1, 2) - R(1, 1) + 1) Then
            A(j) = ay1(x)
            x = x + 1
            j = j + 1
        ElseIf (c1 > c2 Or x > (R(1, 2) - R(1, 1) + 1)) And y <= (R(2, 2) - R(2, 1) + 1) Then
            A(j) = ay2(y)
            y = y + 1
            j = j + 1
        End If
    Next j
    aymerge = A
    End Function

'Sort a 1 dimension array (Quick sort)
Function aysort_quick(ByVal ay As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Boolean, Optional ByVal Title As Boolean) As Variant
If isay(ay, , , 1) = True Then
    aysort_quick = aysort_options(ay, "Quick", desc, str, Title)
Else
    debug_err "aysort_quick", "1d"
End If
End Function

'Quick Sort
Function aysort_quick_c(ay As Variant, Optional str As Boolean) As Variant
Dim ay1 As Variant, ay2 As Variant, P As Variant
Dim R As Variant
If isay(ay, , , 1) = True Then
    R = ayrge(ay)
    If R(2, 1) - R(1, 1) = 0 Then
        aysort_quick_c = ay
    ElseIf R(2, 1) - R(1, 1) = 1 Then
        aysort_quick_c = aysort_bubble_c(ay, str)
    Else
        aySplit_p ay, ay1, ay2, P, str
        ay1 = aysort_quick_c(ay1, str)
        ay2 = aysort_quick_c(ay2, str)
        aysort_quick_c = aymerge_p(ay1, ay2, P)
    End If
Else
    aysort_quick_c = ay
End If
End Function

    'Split the array by pivot
    Function aySplit_p(ay As Variant, ay1 As Variant, ay2 As Variant, P As Variant, ByVal str As Boolean) As Long
    Dim i As Long, j As Long
    Dim A As Variant
    Dim R As Variant
    Dim COM1 As Variant, COM2 As Variant, txt As Variant
    R = ayrge(ay)
    'Find the pivot
    P = ay(R(2, 1))
    A = ay
    i = R(1, 1)
    j = R(2, 1)
    Do Until i >= j
        If str = False Then
            COM1 = CDbl(A(i))
            COM2 = CDbl(A(j))
        ElseIf str = True Then
            COM1 = CStr(A(i))
            COM2 = CStr(A(j))
        End If
        If COM1 > COM2 Then
            If j - i >= 2 Then
                txt = A(j)
                A(j) = A(i)
                A(i) = A(j - 1)
                A(j - 1) = txt
            ElseIf j - i < 2 Then
                txt = A(j)
                A(j) = A(i)
                A(i) = txt
            End If
            j = j - 1
        ElseIf COM1 <= COM2 Then
            i = i + 1
        End If
    Loop
    If j <> R(1, 1) Then
        ReDim ay1(R(1, 1) To j - 1) As Variant
        For i = R(1, 1) To j - 1
            ay1(i) = A(i)
        Next i
    End If
    If j <> R(2, 1) Then
        ReDim ay2(R(1, 1) To (R(1, 1) + R(2, 1) - j - 1)) As Variant
        For i = R(1, 1) To (R(1, 1) + R(2, 1) - j - 1)
            ay2(i) = A(j + i - R(1, 1) + 1)
        Next i
    End If
    End Function
    
    'Merge the array by pivot
    Function aymerge_p(ay1 As Variant, ay2 As Variant, P As Variant) As Variant
    Dim b(1 To 2) As Boolean
    Dim ay(1 To 2) As Variant
    Dim i As Long
    Dim A As Variant
    Dim R As Variant
    b(1) = IsArray(ay1)
    b(2) = IsArray(ay2)
    ay(1) = ay1
    ay(2) = ay2
    If b(1) = True Or b(2) = True Then
        R = ayrges(ay)
    End If
    If b(1) = True And b(2) = True Then
        ReDim A(R(1, 1, 1) To R(1, 2, 1) - R(1, 1, 1) + 1 + R(2, 2, 1) - R(2, 1, 1) + 1 + 1)
        For i = R(1, 1, 1) To R(1, 2, 1)
            A(i) = ay1(i)
        Next i
        A(R(1, 2, 1) + 1) = P
        For i = R(2, 1, 1) To R(2, 2, 1)
            A((R(1, 2, 1) + 1) - R(2, 1, 1) + i + 1) = ay2(i)
        Next i
    ElseIf b(1) = True And b(2) = False Then
        ReDim A(R(1, 1, 1) To R(1, 2, 1) + 1)
        For i = R(1, 1, 1) To R(1, 2, 1)
            A(i) = ay1(i)
        Next i
        A(R(1, 2, 1) + 1) = P
    ElseIf b(1) = False And b(2) = True Then
        ReDim A(R(2, 1, 1) To R(2, 2, 1) + 1)
        For i = R(2, 1, 1) To R(2, 2, 1)
            A(i + 1) = ay2(i)
        Next i
        A(R(2, 1, 1)) = P
    End If
    aymerge_p = A
    End Function
    
'Comb Sort
Function aysort_comb_c(ByVal ay As Variant, ByVal str As Boolean) As Variant
Dim i As Long
Dim R As Variant
Dim c As Long, G As Long
Dim COM1 As Variant, COM2 As Variant, txt As Variant
c = ayct(ay)
G = Floor(ayct(ay) / 1.3)
Do Until G < 1
    For i = 1 To (c - G)
        If str = False Then
            COM1 = CDbl(ay(i))
            COM2 = CDbl(ay(i + G))
        ElseIf str = True Then
            COM1 = CStr(ay(i))
            COM2 = CStr(ay(i + G))
        End If
        If COM1 > COM2 Then
            txt = ay(i)
            ay(i) = ay(i + G)
            ay(i + G) = txt
        End If
    Next i
    G = Floor(G / 1.3)
Loop
aysort_comb_c = ay
End Function

'Sort a 1 dimension array (Comb sort)
Function aysort_comb(ByVal ay As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Boolean, Optional ByVal Title As Boolean) As Variant
If isay(ay, , , 1) = True Then
    aysort_comb = aysort_options(ay, "Comb", desc, str, Title)
Else
    debug_err "aysort_comb", "1d"
End If
End Function

'Array Sort - 1 dimensional
Function aysort(ByVal ay As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Boolean, Optional ByVal Title As Boolean) As Variant
Dim R As Variant
R = ayrge(ay)
If R(2, 1) - R(1, 1) > 500 Then
    aysort = aysort_quick(ay, desc, str, Title)
ElseIf R(2, 1) - R(1, 1) > 80 Then
    aysort = aysort_select(ay, desc, str, Title)
Else
    aysort = aysort_bubble(ay, desc, str, Title)
End If
End Function

'End of sorting

'Unique values of 1 dimension array
Public Function ayuni(ByVal ay As Variant, Optional ByVal Captial As Boolean) As Variant
Dim i As Long, j As Long
Dim R As Variant
Dim c(1 To 2) As Variant
Dim A As Variant, A1 As Variant
If isay(ay, , , 1) = True Then
    R = ayrge(ay)
    A = aysort(ay)
    ReDim A1(R(1, 1) To R(2, 1)) As Variant
    j = R(1, 1)
    For i = R(1, 1) To R(2, 1)
        If i > R(1, 1) Then
            If Captial = True Then
                c(1) = UCase(CStr(A(i)))
                c(2) = UCase(CStr(A(i - 1)))
            Else
                c(1) = CStr(A(i))
                c(2) = CStr(A(i - 1))
            End If
            If c(1) = c(2) Then
                GoTo SKip
            End If
        End If
        A1(j) = A(i)
        j = j + 1
SKip:
    Next i
    ReDim A(1 To (j - R(1, 1)))
    For i = 1 To (j - R(1, 1))
        A(i) = A1(i + R(1, 1) - 1)
    Next i
    ayuni = A
Else
    debug_err "ayuni", "1d"
End If
End Function

'Return ranks of the one-dimensional array (The largest ranks the top) (Same rank when equal)
Public Function ayrank(ByVal ay As Variant) As Variant
Dim i As Long, j As Long
Dim R As Variant
Dim A As Variant, AR As Variant, F As Variant
If isay(ay, , , 1) = True Then
    R = ayrge(ay)
    A = aysort(ay)
    ReDim AR(R(1, 1) To R(2, 1), 1 To 2) As Variant
    ReDim F(R(1, 1) To R(2, 1)) As Variant
    For i = R(1, 1) To R(2, 1)
        AR(i, 1) = A(i)
        AR(i, 2) = R(2, 1) - i + 1
        If i > R(1, 1) Then
            If AR(i, 1) = AR(i - 1, 1) Then
                AR(i - 1, 2) = AR(i, 2)
            End If
        End If
    Next i
    For i = R(1, 1) To R(2, 1)
        For j = R(1, 1) To R(2, 1)
            If ay(i) = AR(j, 1) Then
                F(i) = AR(j, 2)
                Exit For
            End If
        Next j
    Next i
    ayrank = F
Else
    debug_err "ayrank", "1d"
End If
End Function

'Median of 1 dimension array
Public Function aymedian(ByVal ay As Variant) As Variant
Dim c(0 To 2) As Double
Dim A As Variant
If isay(ay, , , 1) = True Then
    A = aysort(ay)
    c(0) = (ayct(ay) + 1) / 2
    If isint(c) = False Then
        c(1) = Floor(c(0))
        c(2) = Ceil(c(0))
        If IsNumeric(A(c(1)) + A(c(2))) = True Then
            aymedian = (CDbl(A(c(1))) + CDbl(A(c(2)))) / 2
        ElseIf IsNumeric(A(c(1)) + A(c(2))) = False Then
            aymedian = A(c(1))
        End If
    Else
        aymedian = A(c(0))
    End If
Else
    debug_err "aymedian", "1d"
End If
End Function

'Frequency table of 1 dimension array
Public Function ayFreq(ByVal ay As Variant) As Variant
Dim i As Long, j As Long
Dim A As Variant, t As Variant
Dim RA As Variant, rf As Variant
If isay(ay, , , 1) = True Then
    RA = ayrge(ay)
    A = aysort(ayuni(ay))
    rf = ayrge(A)
    ReDim t(rf(1, 1) To rf(2, 1), 1 To 2) As Variant
    For j = rf(1, 1) To rf(2, 1)
        t(j, 1) = A(j)
    Next j
    For i = RA(1, 1) To RA(2, 1)
        For j = rf(1, 1) To rf(2, 1)
            If ay(i) = t(j, 1) Then
                t(j, 2) = t(j, 2) + 1
            End If
        Next j
    Next i
    ayFreq = t
Else
    debug_err "ayFreq", "1d"
End If
End Function

'Mode of 1 dimension array
Public Function aymode(ByVal ay As Variant, Optional ByVal rank As Long) As Variant
Dim i As Long, ct As Long
Dim R As Variant
Dim t As Variant
Dim max As Variant, modes() As Variant
If isay(ay, , , 1) = True Then
    t = ayFreq(ay)
    R = ayrge(t)
    max = aymax(ayredim(t, Array(1, 2), 1))
    ReDim modes(1 To R(2, 1)) As Variant
    ct = 0
    For i = R(1, 1) To R(2, 1)
        If t(i, 2) = max Then
            ct = ct + 1
            modes(ct) = t(i, 1)
        End If
    Next i
    If ct = 1 Then
        aymode = modes(ct)
    ElseIf ct > 1 Then
        If rank <= ct And rank >= 1 Then
            aymode = modes(rank)
        ElseIf rank = 0 Then
            debug_err "aymode", , "There are more than one mode for the array, please specify the number of rank of the mode to be chosen."
        ElseIf rank > ct Or rank < 1 Then
            debug_err "aymode", , "The number of rank of the mode is not invalid, please check!"
        End If
    End If
Else
    debug_err "aymode", "1d"
End If
End Function

'Maximum of 1 dimension array (optional: xth largest)
Public Function aymax(ByVal ay As Variant, Optional ByVal rank As Long) As Variant
Dim R As Variant
Dim A As Variant
If isay(ay, , , 1) = True Then
    If rank = 0 Then
        rank = 1
    End If
    R = ayrge(ay)
    A = aysort(ay)
    If rank <= R(2, 1) - R(1, 1) + 1 And rank >= 1 Then
        aymax = A(R(2, 1) - rank + 1)
    ElseIf rank > R(2, 1) - R(1, 1) + 1 Or rank < 1 Then
        debug_err "aymax", , "The rank is out of the range of the array"
    End If
Else
    debug_err "aymax", "1d"
End If
End Function

'Minimum of 1 dimension array (optional: xth smallest)
Public Function aymin(ByVal ay As Variant, Optional ByVal rank As Long) As Variant
Dim R As Variant
Dim A As Variant
If isay(ay, , , 1) = True Then
    If rank = 0 Then
        rank = 1
    End If
    R = ayrge(ay)
    A = aysort(ay)
    If rank <= R(2, 1) - R(1, 1) + 1 And rank >= 1 Then
        aymin = A(R(1, 1) + rank - 1)
    ElseIf rank > R(2, 1) - R(1, 1) + 1 Or rank < 1 Then
        debug_err "aymin", , "The rank is out of the range of the array"
    End If
Else
    debug_err "aymin", "1d"
End If
End Function

'Range of 1 dimension array
Public Function ayrange(ByVal ay As Variant) As Variant
If isay(ay, , , 1) = True Then
    If aynum(ay) = True Then
        ayrange = aymax(ay) - aymin(ay)
    Else
        debug_err "ayrange", "nnumay"
    End If
Else
    debug_err "ayrange", "1d"
End If
End Function

'Percentile of 1 dimension array
Public Function aypercent(ByVal ay As Variant, ByVal percent As Double) As Variant
Dim low As Long, up As Long
Dim ct As Long
Dim c(0 To 3) As Double
Dim A As Variant
If isay(ay, , , 1) = True Then '
    If ispercent(percent) = True Then
        ct = ayct(ay)
        A = aysort(ay)
        c(0) = ((ct - 1) / 100) * percent + 1
        If isint(c(0)) = False Then
            c(3) = c(0) - Int(c(0))
            c(1) = Floor(c(0))
            c(2) = Ceil(c(0))
            If IsNumeric(A(c(1)) + A(c(2))) = True Then
                aypercent = CDbl(A(c(1))) * (1 - c(3)) + CDbl(A(c(2))) * c(3)
            Else
                aypercent = A(c(1))
            End If
        Else
            aypercent = A(c)
        End If
    Else
        debug_err "aypercent", , "Percent is invalid.(not betweeen 0 and 100)"
    End If
Else
    debug_err "aypercent", "1d"
End If
End Function

'interquantile range of 1 dimension array
Public Function ayinrange(ByVal ay As Variant) As Variant
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        ayinrange = aypercent(ay, 75) - aypercent(ay, 25)
    Else
        debug_err "ayinrange", "nnumay"
    End If
Else
    debug_err "ayrange", "1d"
End If
End Function

'Total sum of squares of 1 dimension array (optional: n is the power)
Public Function aysst(ByVal ay As Variant, Optional n As Long) As Double
Dim i As Long
Dim R As Variant
Dim sum As Double
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        If n = 0 Then n = 2
        If n >= 2 Then
            R = ayrge(ay)
            For i = R(1, 1) To R(2, 1)
                sum = sum + (CDbl(ay(i)) ^ n)
            Next i
            aysst = sum
        Else
            debug_err "aysst", , "n must be greater than or equal to 2."
        End If
    ElseIf b = False Then
        debug_err "aysst", "nnumay"
    End If
Else
    debug_err "aysst", "1d"
End If
End Function

'N moments of 1 dimension array
Public Function aymoment(ByVal ay As Variant, ByVal n As Long, Optional ByVal cent As Boolean) As Double
Dim i As Long
Dim R As Variant
Dim sum As Double
Dim mean As Double
Dim ct As Long
If n >= 0 Then
    If isay(ay, , , 1) = True Then
        If isay(ay, , True) = True Then
            R = ayrge(ay)
            ct = ayct(ay)
            If cent = False Then
                For i = R(1, 1) To R(2, 1)
                    sum = sum + (CDbl(ay(i)) ^ n)
                Next i
                aymoment = sum / ct
            Else
                mean = aymean(ay)
                For i = R(1, 1) To R(2, 1)
                    sum = sum + CDbl(((ay(i) - mean) ^ n))
                Next i
                aymoment = sum / ct
            End If
        Else
            debug_err "aymoment", "nnumay"
        End If
    Else
        debug_err "aymoment", "1d"
    End If
Else
    debug_err "aymoment", , "N must be greater than or equal to 0."
End If
End Function

'Variance of 1 dimension array
Public Function ayvar(ByVal ay As Variant, Optional ByVal sample As Boolean) As Double
Dim i As Long
Dim R As Variant
Dim A() As Double
Dim mean As Double
Dim sum As Double
Dim ct As Long
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        R = ayrge(ay)
        mean = ayamean(ay)
        ct = ayct(ay)
        ReDim A(R(1, 1) To R(2, 1)) As Double
        For i = R(1, 1) To R(2, 1)
            A(i) = (CDbl(ay(i)) - mean) ^ 2
        Next i
        sum = aysum(A)
        If sample = False Then
            ayvar = sum / ct
        ElseIf sample = True Then
            ayvar = sum / (ct - 1)
        End If
    Else
        debug_err "ayvar", "nnumay"
    End If
Else
    debug_err "ayvar", "1d"
End If
End Function

'Standard Deviation of 1 dimension array
Public Function aystdev(ByVal ay As Variant, Optional ByVal sample As Boolean) As Double
aystdev = (ayvar(ay, sample)) ^ (1 / 2)
End Function

'Median absolute deviation of 1 dimension array
Public Function aymad(ByVal ay As Variant) As Double
Dim i As Long
Dim med As Variant
Dim R As Variant
Dim A As Variant
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        R = ayrge(ay)
        med = aymedian(ay)
        ReDim A(R(1, 1) To R(2, 1)) As Variant
        For i = R(1, 1) To R(2, 1)
            A(i) = Abs(ay(i) - med)
        Next i
        aymad = aymedian(A)
    Else
        debug_err "aymad", "nnumay"
    End If
Else
    debug_err "aymad", , "1d"
End If
End Function

'Coefficient of variation of 1 dimension array
Public Function aycv(ByVal ay As Variant, Optional ByVal sample As Boolean) As Double
aycv = aystdev(ay, sample) / ayamean(ay)
End Function

'Skewness of 1 dimension array
Public Function ayskew(ByVal ay As Variant, Optional ByVal sample As Boolean) As Double
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        If sample = False Then
            ayskew = aymoment(ay, 3, True) / (aymoment(ay, 2, True) ^ (3 / 2))
        ElseIf sample = True Then
            ayskew = ((ayct(ay) ^ 2) / ((ayct(ay) - 1) * (ayct(ay) - 2))) * (aymoment(ay, 3, True) / (ayvar(ay, True) ^ (3 / 2)))
        End If
    ElseIf b = False Then
        debug_err "ayskew", "nnum"
    End If
Else
    debug_err "ayskew", "1d"
End If
End Function

'Kurtosis of 1 dimensional array
Public Function aykurt(ByVal ay As Variant, Optional ByVal sample As Boolean, Optional ByVal ex As Boolean) As Double
Dim b As Boolean
Dim n As Long
b = aynum(ay)
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        If sample = False Then
            If ex = False Then
                aykurt = aymoment(ay, 4, True) / (aymoment(ay, 2, True) ^ 2)
            ElseIf ex = True Then
                aykurt = (aymoment(ay, 4, True) / (aymoment(ay, 2, True) ^ 2)) - 3
            End If
        ElseIf sample = True Then
            n = ayct(ay)
            If ex = False Then
                aykurt = (((n + 1) * n * (n - 1)) / ((n - 2) * (n - 3))) * ((n * aymoment(ay, 4, True)) / ((n * aymoment(ay, 2, True)) ^ 2))
            ElseIf ex = True Then
                aykurt = ((((n + 1) * n * (n - 1)) / ((n - 2) * (n - 3))) * ((n * aymoment(ay, 4, True)) / ((n * aymoment(ay, 2, True)) ^ 2))) - 3 * ((n - 1) ^ 2 / ((n - 2) * (n - 3)))
            End If
        End If
    Else
        debug_err "aykurt", "nnum"
    End If
Else
    debug_err "ayskew", "1d"
End If
End Function

'Lastest(Maximum) Date of a 1- dimension array
Public Function maxdate(ByRef ay() As Variant) As Variant
Dim low As Long, up As Long
Dim txt As Variant
Dim A() As Variant
Dim A1() As Variant
low = LBound(ay, 1)
up = UBound(ay, 1)
ReDim A(low To up, 1 To 2)
For i = low To up
    A(i, 1) = ay(i)
    If ay(i) <> "" Then
        A(i, 2) = CLng(Right(ay(i), 4)) * 10000 + CLng(mid(ay(i), 4, 2)) * 100 + CLng(Left(ay(i), 2))
    End If
Next i
A1 = aysort2(A, 2)
maxdate = A1(up, 1)
End Function

'Covariance of two 1-dimension arrays
Public Function aycovar(ByVal ay1 As Variant, ay2 As Variant) As Double
Dim i As Long, j As Long
Dim ay(1 To 2) As Variant
Dim m(1 To 2) As Double
Dim n As Long
Dim c As Double
Dim R As Variant
If aynum(ay1) = True And aynum(ay2) = True Then
    ay(1) = ay1
    ay(2) = ay2
    R = ayrges(ay)
    If UBound(R, 3) = 1 Then
        If R(1, 2, 1) = R(2, 2, 1) And R(1, 1, 1) = R(2, 1, 1) Then
            n = ayct(ay1)
            ReDim A(R(1, 1, 1) To R(1, 2, 1)) As Variant
            For i = 1 To 2
                m(i) = aymean(ay(i))
            Next i
            For j = R(1, 1, 1) To R(1, 2, 1)
                c = c + (ay(1)(j) - m(1)) * (ay(2)(j) - m(2))
            Next j
            aycovar = c / n
        Else
            debug_err "aycovar", , "The dimension of the two arrays not match."
        End If
    Else
        debug_err "aycovar", , "At least one of the arrays is not 1 dimensional."
    End If
Else
    debug_err "aycovar", , "At least one of the arrays is not a numeric array."
End If
End Function

'Pearson's product-moment coefficient of two 1-dimension arrays
Public Function aycorr_p(ByVal ay1 As Variant, ay2 As Variant) As Double
Dim i As Long, j As Long
Dim ay(1 To 2) As Variant
Dim R As Variant
If aynum(ay1) = True And aynum(ay2) = True Then
    ay(1) = ay1
    ay(2) = ay2
    R = ayrges(ay)
    If UBound(R, 3) = 1 Then
        If R(1, 2, 1) = R(2, 2, 1) And R(1, 1, 1) = R(2, 1, 1) Then
            aycorr_p = aycovar(ay(1), ay(2)) / (aystdev(ay(1)) * aystdev(ay(2)))
        Else
            debug_err "aycorr_p", , "The dimension of the two arrays not match."
        End If
    ElseIf UBound(R, 3) > 1 Then
        debug_err "aycorr_p", , "At least one of the arrays is not 1 dimensional."
    End If
Else
    debug_err "aycorr_p", , "At least one of the arrays is not a numeric array"
End If
End Function

'Spearman's rank correlation coefficient of two 1-dimension arrays
Public Function aycorr_sr(ByVal ay1 As Variant, ay2 As Variant) As Double
Dim ay(1 To 2) As Variant
If aynum(ay1) = True And aynum(ay2) = True Then
    ay(1) = ayrank(ay1)
    ay(2) = ayrank(ay2)
    aycorr_sr = aycorr_p(ay(1), ay(2))
Else
    debug_err "aycorr_sr", , "At least one of the arrays is not a numeric array"
End If
End Function

'Correlation coefficient of two 1-dimension arrays
Public Function aycorr(ByVal ay1 As Variant, ay2 As Variant) As Double
aycorr = aycorr_p(ay1, ay2)
End Function

'Covariance matrix of n 1-dimension arrays
Public Function aycovars(ByVal ays As Variant) As Variant
Dim i As Long, j As Long
Dim low As Long, up As Long
Dim R As Variant
Dim A As Variant
If isay(ays, , , 1) = True Then
    If isay(ays, True) = True Then
        If aydims_m(ays) = 1 Then
            R = ayrges(ays)
            If UBound(R, 3) = 1 Then
                If R(1, 2, 1) = R(2, 2, 1) And R(1, 1, 1) = R(2, 1, 1) Then
                    low = LBound(R, 1)
                    up = UBound(R, 1)
                    ReDim A(low To up, low To up)
                    For i = low To up
                        For j = low To up
                            A(i, j) = aycovar(ays(i), ays(j))
                        Next j
                    Next i
                    aycovars = A
                Else
                    debug_err "aycovar", , "The dimensions of the arrays not match."
                End If
            Else
                debug_err "aycovar", , "At least one of the arrays is not 1 dimensional."
            End If
        Else
            debug_err "aycovars", , "At least one of the arrays in the jagged array is not one-dimensional"
        End If
    Else
        debug_err "aycovars", "njay"
    End If
Else
    debug_err "aycovars", "1d"
End If
End Function

'Correlation matrix of n 1-dimension arrays
Public Function aycorrs(ByVal ays As Variant) As Variant
Dim i As Long, j As Long
Dim low As Long, up As Long
Dim R As Variant
Dim A As Variant
If isay(ays, , , 1) = True Then
    If isay(ays, True) = True Then
        If aydims_m(ays) = 1 Then
            R = ayrges(ays)
            If UBound(R, 3) = 1 Then
                If R(1, 2, 1) = R(2, 2, 1) And R(1, 1, 1) = R(2, 1, 1) Then
                    low = LBound(R, 1)
                    up = UBound(R, 1)
                    ReDim A(low To up, low To up)
                    For i = low To up
                        For j = low To up
                            A(i, j) = aycorr(ays(i), ays(j))
                        Next j
                    Next i
                    aycorrs = A
                Else
                    debug_err "aycovar", , "The dimensions of the arrays not match."
                End If
            Else
                debug_err "aycovar", , "At least one of the arrays is not 1 dimensional."
            End If
        Else
            debug_err "aycorrs", , "At least one of the arrays in the jagged array is not one-dimensional"
        End If
    Else
        debug_err "aycorrs", "njay"
    End If
Else
    debug_err "aycorrs", "1d"
End If
End Function

'Moving average of 1-dimension array
Public Function ayma(ByVal ay As Variant, ByVal n As Long) As Variant
Dim i As Long
Dim S As Variant
Dim R As Variant
Dim A As Variant
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        R = ayrge(ay)
        ReDim A(R(1, 1) To R(2, 1)) As Variant
        For i = 1 To n
            S = S + ay(R(1, 1) + i - 1)
        Next i
        S = S / n
        A(R(1, 1) + n - 1) = S
        For i = (R(1, 1) + n) To R(2, 1)
            S = S + ay(i) / n - ay(i - n) / n
            A(i) = S
        Next i
        ayma = A
    Else
        debug_err "ayma", "nnumay"
    End If
Else
    debug_err "ayma", "1d"
End If
End Function

'Exponetial Moving average of 1-dimension array
Public Function ayema(ByVal ay As Variant, ByVal A As Double) As Variant
Dim i As Long
Dim S As Variant
Dim ay1 As Variant
Dim R As Variant
Dim A1 As Double
Dim m As Double
If A < 0 Or A > 1 Then Exit Function
If isay(ay, , , 1) = True Then
    If isay(ay, , True) = True Then
        R = ayrge(ay)
        m = aymean(ay)
        ReDim ay1(R(1, 1) To R(2, 1)) As Variant
        For i = R(1, 1) To R(2, 1)
            If i = R(1, 1) Then
                ay1(i) = m
            Else
                ay1(i) = (1 - A) * ay1(i - 1) + A * ay(i)
            End If
        Next i
        ayema = ay1
    Else
        debug_err "ayma", "nnumay"
    End If
Else
    debug_err "ayema", "1d"
End If
End Function

'Standardising 1-dimension array
Public Function Standardise(ByVal ay As Variant, Optional ByVal mean As Variant, Optional ByVal sd As Variant, Optional ByVal sample As Boolean) As Variant
Dim i As Long
Dim m As Variant, S As Variant
Dim R As Variant
Dim A As Variant
If IsMissing(mean) = False Or IsMissing(sd) = False Then
    If IsNumeric(mean) = False Then
        debug_err "Standardise", "nnum - mean"
        Exit Function
    End If
    m = mean
    If IsNumeric(sd) = False Then
        debug_err "Standardise", "nnum - sd"
        Exit Function
    End If
    S = sd
End If
If isay(ay, , True, 1) = True Then
    If IsMissing(mean) = True Then m = aymean(ay)
    If IsMissing(sd) = True Then S = aystdev(ay, sample)
    R = ayrge(ay)
    ReDim A(R(1, 1) To R(2, 1)) As Variant
    For i = R(1, 1) To R(2, 1)
        A(i) = (ay(i) - m) / S
    Next i
    Standardise = A
ElseIf isay(ay) = False Then
    If IsNumeric(ay) = True Then
        If m <> "" And S <> "" Then
            Standardise = (ay - m) / S
        Else
            debug_err "Standardise", , "Please specify the mean and the sd for a single value."
        End If
    ElseIf IsNumeric(ay) = False Then
        GoTo SKip
    End If
Else
SKip:
    debug_err "Standardise", , "ay must be numeric 1-dimensional array or single value."
End If
End Function


'******************************************************************************************************************************************************

'C2!: 2-Dimension array Functions
'Redim Preserve of 2D array
Public Function RedimPre(ByVal ay As Variant, Optional ByVal low1 As Long, Optional ByVal up1 As Long, Optional ByVal low2 As Long, Optional ByVal up2 As Long) As Variant
Dim d As Long
Dim i As Long, j As Long
Dim rge() As Long
Dim TAY() As Variant
If up1 < low1 Or up2 < low2 Then
    debug_err "RedimPre", , "The upper bound must be equal to or greater than the lower bound, please check!"
    Exit Function
End If
d = aydim(ay)
If d = 1 Then
    If low2 = 0 And up2 = 0 Then
        ReDim rge(1 To 2, 1 To 3) As Long '(low or up; Past, new or temp)
        rge(1, 1) = LBound(ay, 1)
        rge(2, 1) = UBound(ay, 1)
        rge(1, 2) = low1
        rge(2, 2) = up1
        For i = 1 To 2
            If rge(i, 2) = 0 Then rge(i, 2) = rge(i, 1)
        Next i
        ReDim TAY(rge(1, 2) To rge(2, 2)) As Variant
        For i = 1 To 2
            If rge(i, 2) <= rge(i, 1) Then
                rge(i, 3) = rge(i, i)
            ElseIf rge(i, 2) > rge(i, 1) Then
                rge(i, 3) = rge(i, ((i Mod 2) + 1))
            End If
        Next i
        For i = rge(1, 3) To rge(2, 3)
            TAY(i) = ay(i)
        Next i
        RedimPre = TAY
    ElseIf low2 <> 0 Or up2 <> 0 Then '(Change the array from 1 dimension to 2 dimension)
        ReDim rge(1 To 2, 1 To 2, 1 To 3) As Long '(low or up; 1 or 2 dimension; Past, new or temp)
        rge(1, 1, 1) = LBound(ay, 1)
        rge(2, 1, 1) = UBound(ay, 1)
        rge(1, 1, 2) = low1
        rge(2, 1, 2) = up1
        rge(1, 2, 2) = low2
        rge(2, 2, 2) = up2
        For i = 1 To 2
            For j = 1 To 2
                If rge(j, i, 2) = 0 Then rge(j, i, 2) = rge(j, i, 1)
            Next j
        Next i
        ReDim TAY(rge(1, 1, 2) To rge(2, 1, 2), rge(1, 2, 2) To rge(2, 2, 2)) As Variant
        For i = 1 To 2
            For j = 1 To 2
                If rge(i, j, 2) <= rge(i, j, 1) Then
                    rge(i, j, 3) = rge(i, j, i)
                ElseIf rge(i, j, 2) > rge(i, j, 1) Then
                    rge(i, j, 3) = rge(i, j, ((i Mod 2) + 1))
                End If
            Next j
        Next i
        For i = rge(1, 1, 3) To rge(2, 1, 3)
            TAY(i, rge(1, 2, 2)) = ay(i)
        Next i
        RedimPre = TAY
    End If
ElseIf d = 2 Then
    ReDim rge(1 To 2, 1 To 2, 1 To 3) As Long '(low or up; 1 or 2 dimension; Past, new or temp)
    For i = 1 To 2
        rge(1, i, 1) = LBound(ay, i)
        rge(2, i, 1) = UBound(ay, i)
    Next i
    rge(1, 1, 2) = low1
    rge(2, 1, 2) = up1
    rge(1, 2, 2) = low2
    rge(2, 2, 2) = up2
    For i = 1 To 2
        For j = 1 To 2
            If rge(j, i, 2) = 0 Then
                rge(j, i, 2) = rge(j, i, 1)
            End If
        Next j
    Next i
    ReDim TAY(rge(1, 1, 2) To rge(2, 1, 2), rge(1, 2, 2) To rge(2, 2, 2)) As Variant
    For i = 1 To 2
        For j = 1 To 2
            If rge(i, j, 2) <= rge(i, j, 1) Then
                rge(i, j, 3) = rge(i, j, i)
            ElseIf rge(i, j, 2) > rge(i, j, 1) Then
                rge(i, j, 3) = rge(i, j, ((i Mod 2) + 1))
            End If
        Next j
    Next i
    For i = rge(1, 1, 3) To rge(2, 1, 3)
        For j = rge(1, 2, 3) To rge(2, 2, 3)
            TAY(i, j) = ay(i, j)
        Next j
    Next i
    RedimPre = TAY
ElseIf d > 2 Then
    debug_err "RedimPre", "2du"
End If
End Function

'Join n 1-dimension arrays into a 2-dimension array
Public Function ayjoin1(ByVal ays As Variant) As Variant
Dim b As Boolean
Dim i As Long, j As Long, k As Long
Dim d As Integer, DA As Integer
Dim low As Long, up As Long
Dim A As Variant
Dim R As Variant, RA As Variant, RM As Variant
d = aydim(ays)
If d = 1 Then
    If isay(ays, True) = True Then
        DA = aydims_m(ays)
        If DA = 1 Then
            RA = ayrges(ays)
            RM = ayrges_b(ays)
            ReDim R(RM(1, 1) To RM(2, 1), LBound(RA, 1) To UBound(RA, 1)) As Variant
            For j = LBound(RA, 1) To UBound(RA, 1)
                For i = RA(j, 1, 1) To RA(j, 2, 1)
                    R(i, j) = ays(j)(i)
                Next i
            Next j
            ayjoin1 = R
        ElseIf DA <> 1 Then
            debug_err "ayjoin1", , "At least one of the arrays in the jagged array is not one-dimensional"
        End If
    Else
        debug_err "ayjoin1", "njay"
    End If
Else
    debug_err "ayjoin1", "1d"
End If
End Function

'lookup of two 2-dimension array (default di = 2)
Public Function aylookup(ByVal ay1 As Variant, ByVal ay2 As Variant, ByVal pos1 As Long, ByVal pos2 As Long, ByVal re As Long, ByVal di As Integer) As Variant
Dim i As Long, j As Long, k As Long
Dim ay(1 To 2) As Variant
Dim d As Variant
Dim R As Variant
Dim A As Variant
ay(1) = ay1
ay(2) = ay2
R = ayrges(ay)
d = aydims_m(ay)
If d = 2 Then
    If di = 0 Then di = 2
    If di = 1 Then
        If pos1 >= R(1, 1, 1) And pos1 <= R(1, 2, 1) And pos2 >= R(2, 1, 1) And pos2 <= R(2, 2, 1) And re >= R(2, 1, 1) And re <= R(2, 2, 1) Then
            For j = R(1, 1, 2) To R(1, 2, 2)
                For i = R(1, 1, 1) To R(1, 2, 1)
                    A(i, j) = ay(1)(i, j)
                Next i
                For k = R(2, 1, 2) To R(2, 2, 2)
                    If ay(1)(pos1, j) = ay(2)(pos2, k) Then
                        A(R(1, 2, 1) + 1, j) = ay2(re, k)
                        GoTo fd
                    End If
                Next k
                A(R(1, 2, 1) + 1, j) = "N/A"
fd:
            Next j
            aylookup = A
        Else
            debug_err "aylookup", , "either pos1, pos2 or re is/are out of the range, please check."
        End If
    ElseIf di = 2 Then
        If pos1 >= R(1, 1, 2) And pos1 <= R(1, 2, 2) And pos2 >= R(2, 1, 2) And pos2 <= R(2, 2, 2) And re >= R(2, 1, 2) And re <= R(2, 2, 2) Then
            ReDim A(R(1, 1, 1) To R(1, 2, 1), R(1, 1, 2) To (R(1, 2, 2) + 1)) As Variant
            For i = R(1, 1, 1) To R(1, 2, 1)
                For j = R(1, 1, 2) To R(1, 2, 2)
                    A(i, j) = ay(1)(i, j)
                Next j
                For k = R(2, 1, 1) To R(2, 2, 1)
                    If ay(1)(i, pos1) = ay(2)(k, pos2) Then
                        A(i, R(1, 2, 2) + 1) = ay2(k, re)
                        GoTo fd2
                    End If
                Next k
                A(i, R(1, 2, 2) + 1) = "N/A"
fd2:
            Next i
            aylookup = A
        Else
            debug_err "aylookup", , "either pos1, pos2 or re is/are out of the range, please check."
        End If
    Else
        debug_err "aylookup", , "di must be either 1 or 2."
    End If
ElseIf d <> 2 Then
    debug_err "aylookup", , "ay1 or/and ay2 is/are not 2 -dimensional."
End If
End Function

'Transpose of a 2-dimension array
Public Function aytranspose(ByVal ay As Variant) As Variant
Dim i As Long, j As Long
Dim d As Long
Dim A As Variant
Dim R As Variant
If aydim(ay) = 2 Then
    R = ayrge(ay)
    ReDim A(R(1, 2) To R(2, 2), R(1, 1) To R(2, 1))
    For i = R(1, 1) To R(2, 1)
        For j = R(1, 2) To R(2, 2)
            A(j, i) = ay(i, j)
        Next j
    Next i
    aytranspose = A
Else
    debug_err "aytranspose", "2du"
End If
End Function

'Sorting algorithm starts

'Sort a 2 dimension array (for options)
Public Function aysort2_options(ByVal ay As Variant, ByVal method As String, ByVal cols As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Variant, Optional ByVal Title As Boolean) As Variant
Dim i As Long, j As Long
Dim R As Variant, RA As Variant
Dim A As Variant, AT As Variant, AB As Variant, AN As Variant
R = ayrge(ay)
If IsMissing(str) = True Then
    str = Not (aynum(ay))
End If
RA = R
If Title = False Then
    AT = ay
ElseIf Title = True Then
    RA(2, 1) = RA(2, 1) - 1
    ReDim AT(RA(1, 1) To RA(2, 1), RA(1, 2) To RA(2, 2)) As Variant
    For i = RA(1, 1) To RA(2, 1)
        For j = RA(1, 2) To RA(2, 2)
            AT(i, j) = ay(i + 1, j)
        Next j
    Next i
End If
'Core Sorting
    AT = aysort2_copt(AT, method, cols, str)
'End of Core sorting
If desc = False Then
    AB = AT
ElseIf desc = True Then
    ReDim AD(RA(1, 1) To RA(2, 1), RA(1, 2) To RA(2, 2)) As Variant
    For i = RA(1, 1) To RA(2, 1)
        For j = RA(1, 2) To RA(2, 2)
            AD(i, j) = AT((RA(2, 1) - i + RA(1, 1)), j)
        Next j
    Next i
    AB = AD
End If
If Title = False Then
    aysort2_options = AB
ElseIf Title = True Then
    ReDim AN(R(1, 1) To R(2, 1), R(1, 2) To R(2, 2)) As Variant
    For j = R(1, 2) To R(2, 2)
        AN(R(1, 1), j) = ay(R(1, 1), j)
        For i = RA(1, 1) To RA(2, 1)
            AN(i + 1, j) = AB(i, j)
        Next i
    Next j
    aysort2_options = AN
End If
End Function

'sort a 2 dimension array (for core sorting)
Public Function aysort2_copt(ByVal AT As Variant, ByVal method As String, ByVal cols As Variant, Optional ByVal str As Variant) As Variant
If method = "Select" Then
    aysort2_copt = aysort2_select_c(AT, cols, str)
ElseIf method = "Bubble" Then
    aysort2_copt = aysort2_bubble_c(AT, cols, str)
ElseIf method = "Merge" Then
    aysort2_copt = aysort2_merge_c(AT, cols, str)
ElseIf method = "Quick" Then
    aysort2_copt = aysort2_quick_c(AT, cols, str)
'ElseIf method = "Comb" Then
'    aysort2_copt = aysort2_comb_c(AT, cols, str)
Else
    debug_err "aysort_options", , "Method not specified, please check."
End If
End Function

'sort a 2 dimension array (Selection sort)
Public Function aysort2_select(ByVal ay As Variant, ByVal cols As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Variant, Optional ByVal Title As Boolean) As Variant
Dim R As Variant
If isay(ay, , , 2) = True Then
    If aysort2_ccols(ay, cols) = True Then
        Exit Function
    End If
    aysort2_select = aysort2_options(ay, "Select", cols, desc, str, Title)
Else
    debug_err "aysort2_select", "2d"
End If
End Function

'checking errors of variable - cols
Function aysort2_ccols(ByVal ay As Variant, ByVal cols As Variant) As Boolean
Dim i As Long
Dim R As Variant
R = ayrge(ay)
If aydim(cols) = 0 Then
    If isint(cols) = False Then
        aysort2_ccols = True
        debug_err "aysort2_select", "npi - cols"
    End If
    If cols < R(1, 2) Or cols > R(2, 2) Then
        aysort2_ccols = True
        debug_err "aysort2_select", , "cols is out of range."
    End If
ElseIf aydim(cols) = 1 Then
    If isay(cols, , True, 1, True, "P") = False Then
        aysort2_ccols = True
        debug_err "aysort2_select", , "Cols is not an array consists only positive integers."
    End If
    For i = LBound(cols, 1) To UBound(cols, 1)
        If cols(i) < R(1, 2) Or cols(i) > R(2, 2) Then
            aysort2_ccols = True
            debug_err "aysort2_select", , "cols is out of range."
            Exit Function
        End If
    Next i
Else
    aysort2_ccols = True
    debug_err "aysort2_select", "1du- - cols"
End If
End Function

'Core Sorting - Compare
Function aysort2_comp(ByVal ay As Variant, ByVal cols As Variant, ByVal x As Long, ByVal y As Long, Optional str As Variant) As Boolean
Dim i As Long
Dim d As Integer
Dim RC As Variant
Dim S As Boolean
Dim COM1 As Variant, COM2 As Variant
d = aydim(cols)
If d = 0 Then
    S = str
    If S = False Then
        COM1 = CDbl(ay(y, cols))
        COM2 = CDbl(ay(x, cols))
    ElseIf S = True Then
        COM1 = UCase(ay(y, cols))
        COM2 = UCase(ay(x, cols))
    End If
    If COM1 < COM2 Then
        aysort2_comp = True
    ElseIf COM1 >= COM2 Then
        aysort2_comp = False
    End If
ElseIf d = 1 Then
    RC = ayrge(cols)
    For i = RC(1, 1) To RC(2, 1)
        S = str(i)
        If S = False Then
            COM1 = CDbl(ay(y, cols(i)))
            COM2 = CDbl(ay(x, cols(i)))
        ElseIf S = True Then
            COM1 = UCase(ay(y, cols(i)))
            COM2 = UCase(ay(x, cols(i)))
        End If
        If COM1 > COM2 Then
            aysort2_comp = False
            Exit Function
        ElseIf COM1 < COM2 Then
            aysort2_comp = True
            Exit Function
        ElseIf COM1 = COM2 Then
        End If
    Next i
End If
End Function

'Core Sorting - Compare (Merge sort)
Function aysort2_compm(ByVal ay1 As Variant, ByVal ay2 As Variant, ByVal cols As Variant, ByVal x As Long, ByVal y As Long, Optional str As Variant) As Integer
Dim i As Long
Dim d As Integer
Dim RC As Variant
Dim COM1 As Variant, COM2 As Variant
d = aydim(cols)
If d = 0 Then
    If str = False Then
        COM1 = CDbl(ay1(x, cols))
        COM2 = CDbl(ay2(y, cols))
    ElseIf str = True Then
        COM1 = UCase(ay1(x, cols))
        COM2 = UCase(ay2(y, cols))
    End If
    If COM1 < COM2 Then
        aysort2_compm = -1
    ElseIf COM1 > COM2 Then
        aysort2_compm = 1
    ElseIf COM1 = COM2 Then
        aysort2_compm = 0
    End If
ElseIf d = 1 Then
    RC = ayrge(cols)
    For i = RC(1, 1) To RC(2, 1)
        If str = False Then
            COM1 = CDbl(ay1(x, cols(i)))
            COM2 = CDbl(ay2(y, cols(i)))
        ElseIf str = True Then
            COM1 = UCase(ay1(x, cols(i)))
            COM2 = UCase(ay2(y, cols(i)))
        End If
        If COM1 < COM2 Then
            aysort2_compm = -1
            Exit Function
        ElseIf COM1 > COM2 Then
            aysort2_compm = 1
            Exit Function
        ElseIf COM1 = COM2 Then
            aysort2_compm = 0
        End If
    Next i
End If
End Function

'Core Sorting - Selection sort
Function aysort2_select_c(ByVal ay As Variant, ByVal cols As Variant, Optional str As Variant) As Variant
Dim i As Long, j As Long
Dim d As Integer
Dim x As Long, y As Long
Dim R As Variant, RC As Variant
Dim txt As Variant
R = ayrge(ay)
For x = R(1, 1) To R(2, 1)
    For y = x To R(2, 1)
        If aysort2_comp(ay, cols, x, y, str) = True Then
            For j = R(1, 2) To R(2, 2)
                txt = ay(x, j)
                ay(x, j) = ay(y, j)
                ay(y, j) = txt
            Next j
        End If
    Next y
Next x
aysort2_select_c = ay
End Function

'Sort a 2 dimension array (Bubble sort)
Public Function aysort2_bubble(ByVal ay As Variant, ByVal cols As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Variant, Optional ByVal Title As Boolean) As Variant
Dim R As Variant
If isay(ay, , , 2) = True Then
    If aysort2_ccols(ay, cols) = True Then
        Exit Function
    End If
    aysort2_bubble = aysort2_options(ay, "Bubble", cols, desc, str, Title)
Else
    debug_err "aysort_bubble", "2d"
End If
End Function

'Core Sorting - bubble sort
Function aysort2_bubble_c(ByVal ay As Variant, ByVal cols As Variant, Optional ByVal str As Variant) As Variant
Dim b As Boolean, E As Boolean
Dim i As Long, j As Long
Dim x As Long
Dim c As Long, c1 As Long
Dim d As Integer
Dim R As Variant, RC As Variant
Dim COM1 As Variant, COM2 As Variant, txt As Variant
R = ayrge(ay)
d = aydim(cols)
Do While E = False
    b = False
    For i = R(1, 1) To (R(2, 1) - 1)
        If aysort2_comp(ay, cols, i, i + 1, str) = True Then
            For j = R(1, 2) To R(2, 2)
                txt = ay(i, j)
                ay(i, j) = ay(i + 1, j)
                ay(i + 1, j) = txt
            Next j
            If b = False Then
                b = True
            End If
        End If
    Next i
    If b = True Then
        E = False
    ElseIf b = False Then
        E = True
    End If
Loop
aysort2_bubble_c = ay
End Function

'Sort a 2 dimension array (Merge sort)
Public Function aysort2_merge(ByVal ay As Variant, ByVal cols As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Variant, Optional ByVal Title As Boolean) As Variant
Dim R As Variant
If isay(ay, , , 2) = True Then
    If aysort2_ccols(ay, cols) = True Then
        Exit Function
    End If
    aysort2_merge = aysort2_options(ay, "Merge", cols, desc, str, Title)
Else
    debug_err "aysort2_merge", "2d"
End If
End Function

'Core Sorting - Merge sort
Function aysort2_merge_c(ay As Variant, ByVal cols As Variant, Optional ByVal str As Variant) As Variant
Dim t As Long
Dim c As Long, c1 As Long
Dim d As Long, dc As Long
Dim i As Long, j As Long
Dim ay1 As Variant, ay2 As Variant
Dim R As Variant
R = ayrge(ay)
dc = aydim(cols)
If R(2, 1) - R(1, 1) = 0 Then
    aysort2_merge_c = ay
ElseIf R(2, 1) - R(1, 1) = 1 Then
    aysort2_merge_c = aysort2_bubble_c(ay, cols, str)
ElseIf R(2, 1) - R(1, 1) > 1 Then
    t = aysplit2(ay, ay1, ay2)
    ay1 = aysort2_merge_c(ay1, cols, str)
    ay2 = aysort2_merge_c(ay2, cols, str)
    If dc = 0 Then
        aysort2_merge_c = aymerge2a(ay, ay1, ay2, cols, str)
    ElseIf dc = 1 Then
        aysort2_merge_c = aymerge2b(ay, ay1, ay2, cols, str)
    End If
End If
End Function
    'Split an array into 2 array (by half)
    Public Function aysplit2(ByVal ay As Variant, ay1 As Variant, ay2 As Variant) As Long
    Dim i As Long, j As Long
    Dim R As Variant
    Dim mid As Long
    R = ayrge(ay)
    mid = Int((R(1, 1) + R(2, 1)) / 2)
    ReDim ay1(1 To mid - R(1, 1) + 1, R(1, 2) To R(2, 2)) As Variant
    ReDim ay2(1 To (R(2, 1) - mid), R(1, 2) To R(2, 2)) As Variant
    For i = R(1, 1) To R(2, 1)
        If i <= mid Then
            For j = R(1, 2) To R(2, 2)
                ay1(i - R(1, 1) + 1, j) = ay(i, j)
            Next j
        ElseIf i > mid Then
            For j = R(1, 2) To R(2, 2)
                ay2(i - mid, j) = ay(i, j)
            Next j
        End If
    Next i
    aysplit2 = (mid - R(1, 1) + 1)
    End Function

    'Merge two array into one array with sorting
    Public Function aymerge2(ay As Variant, ay1 As Variant, ay2 As Variant, cols As Variant, Optional ByVal str As Variant) As Variant
    Dim b As Integer
    Dim c As Long
    Dim i As Long, j As Long, k As Long, x As Long, y As Long
    Dim ayt(1 To 2) As Variant
    Dim R As Variant
    Dim A As Variant
    Dim COM1 As Variant, COM2 As Variant
    On Error Resume Next
    ayt(1) = ay1
    ayt(2) = ay2
    R = ayrges(ayt)
    i = R(1, 2, 1) - R(1, 1, 1) + 1
    j = R(2, 2, 1) - R(2, 1, 1) + 1
    If j >= R(2, 1, 1) Then
        i = i + (R(2, 2, 1) - R(2, 1, 1) + 1)
    End If
    ReDim A(R(1, 1, 1) To i, R(1, 1, 2) To R(1, 2, 2)) As Variant
    x = R(1, 1, 1)
    y = R(2, 1, 1)
    For j = R(1, 1, 1) To i Step 0
        If x > (R(1, 2, 1) - R(1, 1, 1) + 1) And y > (R(2, 2, 1) - R(2, 1, 1) + 1) Then
            aymerge2 = A
            Exit Function
        End If
        b = aysort2_compm(ay1, ay2, cols, x, y, str)
        If b = 0 And (x <= (R(1, 2, 1) - R(1, 1, 1) + 1) And y <= (R(2, 2, 1) - R(2, 1, 1) + 1)) Then
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(j, k) = ay1(x, k)
                A(j + 1, k) = ay2(y, k)
            Next k
            y = y + 1
            x = x + 1
            j = j + 2
        ElseIf (b = -1 Or y > (R(2, 2, 1) - R(2, 1, 1) + 1)) And x <= (R(1, 2, 1) - R(1, 1, 1) + 1) Then
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(j, k) = ay1(x, k)
            Next k
            x = x + 1
            j = j + 1
        ElseIf (b = 1 Or x > (R(1, 2, 1) - R(1, 1, 1) + 1)) And y <= (R(2, 2, 1) - R(2, 1, 1) + 1) Then
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(j, k) = ay2(y, k)
            Next k
            y = y + 1
            j = j + 1
        End If
    Next j
    aymerge2 = A
    End Function
    
    'Merge two array into one array with sorting
    Public Function aymerge2b(ay As Variant, ay1 As Variant, ay2 As Variant, cols As Variant, Optional ByVal str As Variant) As Variant
    Dim b As Integer
    Dim c As Long
    Dim i As Long, j As Long, k As Long, x As Long, y As Long, u As Long
    Dim ayt(1 To 2) As Variant
    Dim R As Variant
    Dim A As Variant
    Dim S As Boolean
    Dim COM1 As Variant, COM2 As Variant
    On Error Resume Next
    ayt(1) = ay1
    ayt(2) = ay2
    R = ayrges(ayt)
    i = R(1, 2, 1) - R(1, 1, 1) + 1
    j = R(2, 2, 1) - R(2, 1, 1) + 1
    If j >= R(2, 1, 1) Then
        i = i + (R(2, 2, 1) - R(2, 1, 1) + 1)
    End If
    ReDim A(R(1, 1, 1) To i, R(1, 1, 2) To R(1, 2, 2)) As Variant
    x = R(1, 1, 1)
    y = R(2, 1, 1)
    For j = R(1, 1, 1) To i Step 0
        If x > (R(1, 2, 1) - R(1, 1, 1) + 1) And y > (R(2, 2, 1) - R(2, 1, 1) + 1) Then
            aymerge2b = A
            Exit Function
        End If
        For u = LBound(cols, 1) To UBound(cols, 1)
            If IsArray(str) = False Then
                S = str
            ElseIf IsArray(str) = True Then
                S = str(u)
            End If
            If S = False Then
                COM1 = CDbl(ay1(x, cols(u)))
                COM2 = CDbl(ay2(y, cols(u)))
            ElseIf S = True Then
                COM1 = UCase(ay1(x, cols(u)))
                COM2 = UCase(ay2(y, cols(u)))
            End If
            If COM1 < COM2 Then
                b = -1
                Exit For
            ElseIf COM1 > COM2 Then
                b = 1
                Exit For
            ElseIf COM1 = COM2 Then
            End If
        Next u
        If b = 0 And (x <= (R(1, 2, 1) - R(1, 1, 1) + 1) And y <= (R(2, 2, 1) - R(2, 1, 1) + 1)) Then
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(j, k) = ay1(x, k)
                A(j + 1, k) = ay2(y, k)
            Next k
            y = y + 1
            x = x + 1
            j = j + 2
        ElseIf (b = -1 Or y > (R(2, 2, 1) - R(2, 1, 1) + 1)) And x <= (R(1, 2, 1) - R(1, 1, 1) + 1) Then
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(j, k) = ay1(x, k)
            Next k
            x = x + 1
            j = j + 1
        ElseIf (b = 1 Or x > (R(1, 2, 1) - R(1, 1, 1) + 1)) And y <= (R(2, 2, 1) - R(2, 1, 1) + 1) Then
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(j, k) = ay2(y, k)
            Next k
            y = y + 1
            j = j + 1
        End If
    Next j
    aymerge2b = A
    End Function
    
    'Merge two array into one array with sorting
    Public Function aymerge2a(ay As Variant, ay1 As Variant, ay2 As Variant, cols As Variant, Optional ByVal str As Variant) As Variant
    Dim b As Integer
    Dim c As Long
    Dim i As Long, j As Long, k As Long, x As Long, y As Long
    Dim ayt(1 To 2) As Variant
    Dim R As Variant
    Dim A As Variant
    Dim COM1 As Variant, COM2 As Variant
    On Error Resume Next
    ayt(1) = ay1
    ayt(2) = ay2
    R = ayrges(ayt)
    i = R(1, 2, 1) - R(1, 1, 1) + 1
    j = R(2, 2, 1) - R(2, 1, 1) + 1
    If j >= R(2, 1, 1) Then
        i = i + (R(2, 2, 1) - R(2, 1, 1) + 1)
    End If
    ReDim A(R(1, 1, 1) To i, R(1, 1, 2) To R(1, 2, 2)) As Variant
    x = R(1, 1, 1)
    y = R(2, 1, 1)
    For j = R(1, 1, 1) To i Step 0
        If x > (R(1, 2, 1) - R(1, 1, 1) + 1) And y > (R(2, 2, 1) - R(2, 1, 1) + 1) Then
            aymerge2a = A
            Exit Function
        End If
        If str = False Then
            COM1 = CDbl(ay1(x, cols))
            COM2 = CDbl(ay2(y, cols))
        ElseIf str = True Then
            COM1 = UCase(ay1(x, cols))
            COM2 = UCase(ay2(y, cols))
        End If
        If COM1 < COM2 Then
            b = -1
        ElseIf COM1 > COM2 Then
            b = 1
        ElseIf COM1 = COM2 Then
            b = 0
        End If
        If b = 0 And (x <= (R(1, 2, 1) - R(1, 1, 1) + 1) And y <= (R(2, 2, 1) - R(2, 1, 1) + 1)) Then
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(j, k) = ay1(x, k)
                A(j + 1, k) = ay2(y, k)
            Next k
            y = y + 1
            x = x + 1
            j = j + 2
        ElseIf (b = -1 Or y > (R(2, 2, 1) - R(2, 1, 1) + 1)) And x <= (R(1, 2, 1) - R(1, 1, 1) + 1) Then
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(j, k) = ay1(x, k)
            Next k
            x = x + 1
            j = j + 1
        ElseIf (b = 1 Or x > (R(1, 2, 1) - R(1, 1, 1) + 1)) And y <= (R(2, 2, 1) - R(2, 1, 1) + 1) Then
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(j, k) = ay2(y, k)
            Next k
            y = y + 1
            j = j + 1
        End If
    Next j
    aymerge2a = A
    End Function
    
'Sort a 2 dimension array (Quick sort)
Public Function aysort2_quick(ByVal ay As Variant, ByVal cols As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Variant, Optional ByVal Title As Boolean) As Variant
Dim R As Variant
If isay(ay, , , 2) = True Then
    If aysort2_ccols(ay, cols) = True Then
        Exit Function
    End If
    aysort2_quick = aysort2_options(ay, "Quick", cols, desc, str, Title)
Else
    debug_err "aysort2_merge", "2d"
End If
End Function
    
'Quick Sort
Function aysort2_quick_c(ByVal ay As Variant, ByVal cols As Variant, Optional ByVal str As Variant) As Variant
Dim ay1 As Variant, ay2 As Variant, P As Variant
Dim R As Variant
If isay(ay, , , 2) = True Then
    R = ayrge(ay)
    If R(2, 1) - R(1, 1) = 0 Then
        aysort2_quick_c = ay
    ElseIf R(2, 1) - R(1, 1) = 1 Then
        aysort2_quick_c = aysort2_bubble_c(ay, cols, str)
    Else
        If isay(cols) = False Then
            aySplit2_pa ay, ay1, ay2, cols, P, str
        ElseIf isay(cols, , True, 1, True) = True Then
            aySplit2_pb ay, ay1, ay2, cols, P, str
        End If
        ay1 = aysort2_quick_c(ay1, cols, str)
        ay2 = aysort2_quick_c(ay2, cols, str)
        aysort2_quick_c = aymerge2_p(ay1, ay2, P)
    End If
Else
    aysort2_quick_c = ay
End If
End Function

    'Split the array by pivot
    Function aySplit2_pa(ay As Variant, ay1 As Variant, ay2 As Variant, ByVal cols As Variant, P As Variant, ByVal str As Variant) As Long
    Dim i As Long, j As Long, k As Long
    Dim A As Variant
    Dim R As Variant
    Dim COM1 As Variant, COM2 As Variant, txt As Variant
    R = ayrge(ay)
    'Find the pivot
    ReDim P(R(1, 2) To R(2, 2)) As Variant
    For k = R(1, 2) To R(2, 2)
        P(k) = ay(R(2, 1), k)
    Next k
    A = ay
    i = R(1, 1)
    j = R(2, 1)
    Do Until i >= j
        If str = False Then
            COM1 = CDbl(A(i, cols))
            COM2 = CDbl(A(j, cols))
        ElseIf str = True Then
            COM1 = CStr(A(i, cols))
            COM2 = CStr(A(j, cols))
        End If
        If COM1 > COM2 Then
            If j - i >= 2 Then
                For k = R(1, 2) To R(2, 2)
                    txt = A(j, k)
                    A(j, k) = A(i, k)
                    A(i, k) = A(j - 1, k)
                    A(j - 1, k) = txt
                Next k
            ElseIf j - i < 2 Then
                For k = R(1, 2) To R(2, 2)
                    txt = A(j, k)
                    A(j, k) = A(i, k)
                    A(i, k) = txt
                Next k
            End If
            j = j - 1
        ElseIf COM1 <= COM2 Then
            i = i + 1
        End If
    Loop
    If j <> R(1, 1) Then
        ReDim ay1(R(1, 1) To j - 1, R(1, 2) To R(2, 2)) As Variant
        For i = R(1, 1) To j - 1
            For k = R(1, 2) To R(2, 2)
                ay1(i, k) = A(i, k)
            Next k
        Next i
    End If
    If j <> R(2, 1) Then
        ReDim ay2(R(1, 1) To (R(1, 1) + R(2, 1) - j - 1), R(1, 2) To R(2, 2)) As Variant
        For i = R(1, 1) To (R(1, 1) + R(2, 1) - j - 1)
            For k = R(1, 2) To R(2, 2)
                ay2(i, k) = A(j + i - R(1, 1) + 1, k)
            Next k
        Next i
    End If
    End Function
    
    'Split the array by pivot
    Function aySplit2_pb(ay As Variant, ay1 As Variant, ay2 As Variant, ByVal cols As Variant, P As Variant, ByVal str As Variant) As Long
    Dim b As Boolean
    Dim i As Long, j As Long, k As Long, u As Long
    Dim A As Variant
    Dim R As Variant
    Dim COM1 As Variant, COM2 As Variant, txt As Variant
    R = ayrge(ay)
    'Find the pivot
    ReDim P(R(1, 2) To R(2, 2)) As Variant
    For k = R(1, 2) To R(2, 2)
        P(k) = ay(R(2, 1), k)
    Next k
    A = ay
    i = R(1, 1)
    j = R(2, 1)
    Do Until i >= j
        For u = LBound(cols, 1) To UBound(cols, 1)
            If str = False Then
                COM1 = CDbl(A(i, cols(u)))
                COM2 = CDbl(A(j, cols(u)))
            ElseIf str = True Then
                COM1 = UCase(A(i, cols(u)))
                COM2 = UCase(A(j, cols(u)))
            End If
            If COM1 > COM2 Then
                b = True
                Exit For
            ElseIf COM1 < COM2 Then
                b = False
                Exit For
            ElseIf COM1 = COM2 Then
            End If
        Next u
        If b = True Then
            If j - i >= 2 Then
                For k = R(1, 2) To R(2, 2)
                    txt = A(j, k)
                    A(j, k) = A(i, k)
                    A(i, k) = A(j - 1, k)
                    A(j - 1, k) = txt
                Next k
            ElseIf j - i < 2 Then
                For k = R(1, 2) To R(2, 2)
                    txt = A(j, k)
                    A(j, k) = A(i, k)
                    A(i, k) = txt
                Next k
            End If
            j = j - 1
        ElseIf b = False Then
            i = i + 1
        End If
    Loop
    If j <> R(1, 1) Then
        ReDim ay1(R(1, 1) To j - 1, R(1, 2) To R(2, 2)) As Variant
        For i = R(1, 1) To j - 1
            For k = R(1, 2) To R(2, 2)
                ay1(i, k) = A(i, k)
            Next k
        Next i
    End If
    If j <> R(2, 1) Then
        ReDim ay2(R(1, 1) To (R(1, 1) + R(2, 1) - j - 1), R(1, 2) To R(2, 2)) As Variant
        For i = R(1, 1) To (R(1, 1) + R(2, 1) - j - 1)
            For k = R(1, 2) To R(2, 2)
                ay2(i, k) = A(j + i - R(1, 1) + 1, k)
            Next k
        Next i
    End If
    End Function
    
    'Merge the array by pivot
    Function aymerge2_p(ay1 As Variant, ay2 As Variant, P As Variant) As Variant
    Dim b(1 To 2) As Boolean
    Dim ay(1 To 2) As Variant
    Dim i As Long
    Dim A As Variant
    Dim R As Variant
    b(1) = IsArray(ay1)
    b(2) = IsArray(ay2)
    ay(1) = ay1
    ay(2) = ay2
    If b(1) = True Or b(2) = True Then
        R = ayrges(ay)
    End If
    If b(1) = True And b(2) = True Then
        ReDim A(R(1, 1, 1) To R(1, 2, 1) - R(1, 1, 1) + 1 + R(2, 2, 1) - R(2, 1, 1) + 1 + 1, R(1, 1, 2) To R(1, 2, 2))
        For i = R(1, 1, 1) To R(1, 2, 1)
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(i, k) = ay1(i, k)
            Next k
        Next i
        For k = R(1, 1, 2) To R(1, 2, 2)
            A(R(1, 2, 1) + 1, k) = P(k)
        Next k
        For i = R(2, 1, 1) To R(2, 2, 1)
            For k = R(1, 1, 2) To R(1, 2, 2)
                A((R(1, 2, 1) + 1) - R(2, 1, 1) + i + 1, k) = ay2(i, k)
            Next k
        Next i
    ElseIf b(1) = True And b(2) = False Then
        ReDim A(R(1, 1, 1) To R(1, 2, 1) + 1, R(1, 1, 2) To R(1, 2, 2))
        For i = R(1, 1, 1) To R(1, 2, 1)
            For k = R(1, 1, 2) To R(1, 2, 2)
                A(i, k) = ay1(i, k)
            Next k
        Next i
        For k = R(1, 1, 2) To R(1, 2, 2)
            A(R(1, 2, 1) + 1, k) = P(k)
        Next k
    ElseIf b(1) = False And b(2) = True Then
        ReDim A(R(2, 1, 1) To R(2, 2, 1) + 1, R(2, 1, 2) To R(2, 2, 2))
        For i = R(2, 1, 1) To R(2, 2, 1)
            For k = R(2, 1, 2) To R(2, 2, 2)
                A(i + 1, k) = ay2(i, k)
            Next k
        Next i
        For k = R(2, 1, 2) To R(2, 2, 2)
            A(R(2, 1, 1), k) = P(k)
        Next k
    End If
    aymerge2_p = A
    End Function
    
'Sort a 2-dimension array (by 2-dimension cols)
Function aysort2(ByVal ay As Variant, ByVal cols As Variant, Optional ByVal desc As Boolean, Optional ByVal str As Variant, Optional ByVal Title As Boolean) As Variant
aysort2 = aysort2_merge(ay, cols, desc, str, Title)
End Function

'End of the sorting

'Comparison the boundary of 2-dimension array (Return indicators, option dir: 0: equal, 1: greater, -1: less, 2: Greater or equal, -2: Less or equal)
Function aybcomp2(ByVal ay1 As Variant, ay2 As Variant, Optional dir As Variant) As Variant
Dim i As Long, j As Long
Dim R As Variant, r1 As Variant, r2 As Variant
If isay(ay1, , , 2) = False Or isay(ay2, , , 2) = False Then
    debug_err "aybcomp2", , "either ay1 or ay2 is not array."
    Exit Function
End If
r1 = ayrge(ay1)
r2 = ayrge(ay2)
ReDim R(1 To 2, 1 To 2) As Variant
For i = 1 To 2
    For j = 1 To 2
        If r1(i, j) > r2(i, j) Then
            R(i, j) = 1
        ElseIf r1(i, j) < r2(i, j) Then
            R(i, j) = -1
        ElseIf r1(i, j) = r2(i, j) Then
            R(i, j) = 0
        End If
    Next j
Next i
If dir = "" Then
    aybcomp2 = R
ElseIf dir = 0 Then
    If R(1, 1) = 0 And R(2, 1) = 0 And R(1, 2) = 0 And R(2, 2) = 0 Then
        aybcomp2 = True
    End If
ElseIf dir = 1 Then
    If R(1, 1) = 1 And R(2, 1) = 1 And R(1, 2) = 1 And R(2, 2) = 1 Then
        aybcomp2 = True
    End If
ElseIf dir = -1 Then
    If R(1, 1) = -1 And R(2, 1) = -1 And R(1, 2) = -1 And R(2, 2) = -1 Then
        aybcomp2 = True
    End If
ElseIf dir = 2 Then
    If (R(1, 1) = 0 Or R(1, 1) = 1) And (R(2, 1) = 0 Or R(2, 1) = 1) And (R(1, 2) = 0 Or R(1, 2) = 1) And (R(2, 2) = 0 Or R(2, 2) = 1) Then
        aybcomp2 = True
    End If
ElseIf dir = -2 Then
    If (R(1, 1) = 0 Or R(1, 1) = -1) And (R(2, 1) = 0 Or R(2, 1) = -1) And (R(1, 2) = 0 Or R(1, 2) = -1) And (R(2, 2) = 0 Or R(2, 2) = -1) Then
        aybcomp2 = True
    End If
Else
    debug_err "aybcomp2", , "error in value of dir, please check."
End If
End Function

'Write a 2-dimension array into another 2-dimension array
Function aywrite2(ByVal ay_f As Variant, ByVal ay_t As Variant) As Variant
Dim i As Long, j As Long
Dim rf As Variant, rt As Variant
If aybcomp2(ay_f, ay_t, 2) = False Then
    debug_err "aywrite2", , "The boundary of ay_t exceeds ay_f, please check."
    Exit Function
End If
rf = ayrge(ay_f)
rt = ayrge(ay_t)
For i = rt(1, 1) To rt(2, 1)
    For j = rt(2, 1) To rt(2, 2)
        ay_f(i, j) = ay_t(i, j)
    Next j
Next i
aywrite2 = ay_f
End Function

'C2a: Matrix Functions

'Is the array a numeric matrix
Public Function ismx(ay As Variant, Optional num As Boolean) As Boolean
Dim d As Long
Dim n As Boolean
d = aydim(ay)
If d = 2 Then
    If num = False Then
        ismx = True
    ElseIf num = True Then
        n = isay(ay, , True)
        If n = False Then
            ismx = False
        ElseIf n = True Then
            ismx = True
        End If
    End If
End If
End Function

'Dimension of a matrix
Public Function mxrge(ByVal mx As Variant) As Variant
Dim b As Boolean
Dim R As Variant
b = ismx(mx)
If b = True Then
    R = ayrge(mx)
    mxrge = R
ElseIf b = False Then
    debug_err "mxrge", "nmx"
End If
End Function

'Dimensions of matrixs
Public Function mxrges(ByVal mxs As Variant) As Variant
Dim b As Boolean
Dim R As Variant
b = isay(mxs, True)
If b = True Then
    R = ayrges(mxs)
    If UBound(R, 3) = 2 Then
        mxrges = R
    Else
        debug_err "mxrges", , "At least one of the arrays in mx is not matrix."
    End If
ElseIf b = False Then
    debug_err "mxrges", , "mxs is not a jagged array."
End If
End Function

'Comparison of the demension of 2 matrixs
Public Function mxcomp(ByVal mx1 As Variant, ByVal mx2 As Variant) As Boolean
Dim i As Long, j As Long
Dim d(1 To 2) As Long
Dim R(1 To 2) As Variant

d(1) = aydim(mx1)
d(2) = aydim(mx2)

If d(1) <> 2 Or d(2) <> 2 Then
    debug_err "mxcomp", , "At least one of the arrays in mx is not matrix."
End If
R(1) = ayrge(mx1)
R(2) = ayrge(mx2)

For i = 1 To 2
    For j = 1 To 2
        If R(1)(i, j) <> R(2)(i, j) Then
            mxcomp = False
            Exit Function
        End If
    Next j
Next i
mxcomp = True
End Function


'Addition of two matrixs
Public Function mxadd(ByVal mx1 As Variant, mx2 As Variant) As Variant
Dim b1 As Boolean, b2 As Boolean
Dim i As Long, j As Long
Dim mx(1 To 2) As Variant
Dim R(1 To 2) As Variant 'nth matrix, lower or upper bound, 1 or 2 dimension
Dim A As Variant
b1 = ismx(mx1, True)
b2 = ismx(mx2, True)
If b1 = True And b2 = True Then
    mx(1) = mx1
    mx(2) = mx2
    R(1) = ayrges(mx)
    If R(1, 1, 1) = R(2, 1, 1) And R(1, 1, 2) = R(2, 1, 2) And R(1, 2, 1) - R(1, 1, 1) = R(2, 2, 1) - R(2, 1, 1) And R(1, 2, 2) - R(1, 1, 2) = R(2, 2, 2) - R(2, 1, 2) Then
        ReDim A(R(1, 1, 1) To R(1, 2, 1), R(1, 1, 2) To R(1, 2, 2)) As Variant
        For i = R(1, 1, 1) To R(1, 2, 1)
            For j = R(1, 1, 2) To R(1, 2, 2)
                A(i, j) = mx(1)(i, j) + mx(2)(i, j)
            Next j
        Next i
        mxadd = A
    Else
        debug_err "mxadd", , "The dimensions of mx1 and mx2 do not match."
    End If
Else
    debug_err "mxadd", , "mx1 or/and mx2 is(are) not (a) matrix(s)"
End If
End Function

'Subtraction of two matrixs
Public Function mxsub(ByVal mx1 As Variant, mx2 As Variant) As Variant
Dim b1 As Boolean, b2 As Boolean
Dim i As Long, j As Long
Dim mx(1 To 2) As Variant
Dim R As Variant 'nth matrix, lower or upper bound, 1 or 2 dimension
Dim A As Variant
b1 = nmx(mx1)
b2 = nmx(mx2)
If b1 = True And b2 = True Then
    mx(1) = mx1
    mx(2) = mx2
    R = ayrges(mx)
    If R(1, 1, 1) = R(2, 1, 1) And R(1, 1, 2) = R(2, 1, 2) And R(1, 2, 1) - R(1, 1, 1) = R(2, 2, 1) - R(2, 1, 1) And R(1, 2, 2) - R(1, 1, 2) = R(2, 2, 2) - R(2, 1, 2) Then
        ReDim A(R(1, 1, 1) To R(1, 2, 1), R(1, 1, 2) To R(1, 2, 2)) As Variant
        For i = R(1, 1, 1) To R(1, 2, 1)
            For j = R(1, 1, 2) To R(1, 2, 2)
                A(i, j) = mx(1)(i, j) - mx(2)(i, j)
            Next j
        Next i
        mxsub = A
    Else
        debug_err "mxsub", , "The dimensions of mx1 and mx2 do not match."
    End If
Else
    debug_err "mxsub", , "mx1 or/and mx2 is(are) not (a) matrix(s)"
End If
End Function

'scalar multiplication of two matrix
Public Function mxmult_s(ByVal mx As Variant, S As Double) As Variant
Dim b As Boolean
Dim i As Long, j As Long
Dim R As Variant, A As Variant
b = ismx(mx, True)
If b = True Then
    R = mxrge(mx)
    ReDim A(R(1, 1) To R(2, 1), R(1, 2) To R(2, 2)) As Variant
    For i = R(1, 1) To R(2, 1)
        For j = R(1, 2) To R(2, 2)
            A(i, j) = S * mx(i, j)
        Next j
    Next i
    mxmult_s = A
ElseIf b = False Then
    debug_err "mxrge", "nmx"
End If
End Function

'Multiplication of two matrix
Public Function mxmult(ByVal mx1 As Variant, mx2 As Variant) As Variant
Dim mx(1 To 2) As Variant
Dim i As Long, j As Long, k As Long
Dim d As Variant
Dim R As Variant
Dim A As Variant
Dim t As Variant
mx(1) = mx1
mx(2) = mx2
d = aydims(mx)
If d(1) = 2 And d(2) = 2 Then
    R = mxrges(mx)
    If (R(1, 2, 2) - R(1, 1, 2)) = (R(2, 2, 1) - R(2, 1, 1)) And R(1, 1, 2) = R(2, 1, 1) Then
        ReDim A(R(1, 1, 1) To R(1, 2, 1), R(2, 1, 2) To R(2, 2, 2))
        For i = R(1, 1, 1) To R(1, 2, 1)
            For j = R(2, 1, 2) To R(2, 2, 2)
                t = 0
                For k = R(1, 1, 2) To R(1, 2, 2)
                    t = t + (mx(1)(i, k) * mx(2)(k, j))
                Next k
                A(i, j) = t
            Next j
        Next i
        mxmult = A
    Else
        debug_err "mxmult", , "n of mx1 not equal to m of mx2."
    End If
Else
    debug_err "mxmult", , "mx1 or/and mx2 is(are) not (a) matrix(s)"
End If
End Function

'Dimension of a square matrix ('if not a square matrix returns 0)
Public Function sqmx(ByVal mx As Variant) As Long
Dim n As Boolean
Dim R As Variant
n = ismx(mx, True)
If n = True Then
    R = mxrge(mx)
    If (R(2, 1) - R(1, 1)) = (R(2, 2) - R(1, 2)) And R(1, 1) = R(1, 2) Then
        sqmx = R(2, 1) - R(1, 1) + 1
    Else
        sqmx = 0
    End If
ElseIf n = False Then
    debug_err "sqmx", "nmx"
End If
End Function

'Determinant of a square matrix (method 1) (start from [1, 1])
Function mxdet1(ByVal mx As Variant) As Double
Dim det As Double
Dim i As Long, j As Long, k As Long, k1 As Long
Dim d As Long, n As Long
Dim t As Double
n = sqmx(mx)
If n > 0 Then
    det = 1
    For k = 1 To n
        If mx(k, k) = 0 Then
            j = k
            Do
                b = True
                If mx(k, j) = 0 Then
                    If j = n Then
                        det = 0
                        Exit Function
                    End If
                    b = False
                    j = j + 1
                ElseIf mx(k, j) <> 0 Then
                    For i = k To n
                        t = mx(i, j)
                        mx(i, j) = mx(i, k)
                        mx(i, k) = t
                    Next i
                    det = -det
                End If
            Loop While b = False
        End If
            det = det * mx(k, k)
        If n - k > 0 Then
            k1 = k + 1
            For i = k1 To n
                For j = k1 To n
                    mx(i, j) = mx(i, j) - (mx(i, k) * mx(k, j) / mx(k, k))
                Next j
            Next i
        End If
    Next k
    mxdet1 = det
ElseIf n = 0 Then
    debug_err "mxdet1", "nsqmx"
End If
End Function

'Minor of a matrix
Function submx(mx As Variant, R As Long, c As Long) As Variant
Dim d As Long
Dim i As Long, j As Long, x As Long, y As Long
Dim RM As Variant
Dim A As Variant
d = aydim(mx)
If d = 2 Then
    RM = ayrge(mx)
    ReDim A(RM(1, 1) To RM(2, 1) - 1, RM(1, 2) To RM(2, 2) - 1)
    For i = RM(1, 1) To RM(2, 1)
        For j = RM(1, 2) To RM(2, 2)
            If i < R Then
                x = i
            ElseIf i > R Then
                x = i - 1
            End If
            If j < c Then
                y = j
            ElseIf j > c Then
                y = j - 1
            End If
            If i <> R And j <> c Then
                A(x, y) = mx(i, j)
            End If
        Next j
    Next i
    submx = A
ElseIf d <> 2 Then
    debug_err "submx", "2d"
End If
End Function

'Determinant of a square matrix (method 2)
Public Function mxdet2(mx As Variant) As Double
Dim det As Double
Dim num As Boolean
Dim i As Long, j As Long
Dim n As Long
Dim R As Variant
n = sqmx(mx)
If n > 0 Then
    R = mxrge(mx)
    If R(2, 1) - R(1, 1) = 1 Then
        det = mx(R(1, 1), R(1, 2)) * mx(R(2, 1), R(2, 2)) - mx(R(1, 1), R(2, 2)) * mx(R(2, 1), R(1, 2))
    ElseIf R(2, 1) - R(1, 1) > 1 Then
        For i = R(1, 1) To R(2, 1)
            det = det + ((-1) ^ (i - 1)) * mx(i, 1) * mxdet2(submx(mx, i, 1))
        Next i
    End If
    mxdet2 = det
ElseIf n = 0 Then
    debug_err "mxdet2", "nsqmx"
End If
End Function

'Determinant of a square matrix
Function mxdet(ay As Variant) As Double
mxdet = mxdet2(ay)
End Function

'Inverse of a square matrix
Function mxinverse(mx As Variant) As Variant
Dim A As Variant
Dim i As Long, j As Long
Dim n As Long
Dim det As Double
Dim R As Variant
n = sqmx(mx)
If n > 0 Then
    R = mxrge(mx)
    det = mxdet(mx)
    ReDim A(R(1, 1) To R(2, 1), R(1, 2) To R(2, 2))
    For i = R(1, 1) To R(2, 1)
        For j = R(1, 2) To R(2, 2)
            A(i, j) = (((-1) ^ (i + j)) * mxdet(submx(mx, i, j))) / det
        Next j
    Next i
    mxinverse = aytranspose(A)
ElseIf n = 0 Then
    debug_err "mxinverse", "nsqmx"
End If
End Function

'Solve linear equation by Cramer 's Rule
Function cramer(ByVal ay As Variant) As Variant
Dim d As Integer
Dim i As Long, j As Long, k As Long
Dim R As Variant
Dim A As Variant, AA As Variant
Dim x As Variant
Dim de As Variant
Dim Des As Variant
If aydim(ay) <> 2 Then
    debug_err "cramer", "2d"
End If
R = ayrge(ay)
If ((R(2, 1) - R(1, 1)) + 1) <> (R(2, 2) - R(1, 2)) Then
    debug_err "cramer", , "The array specified is not correct."
    Exit Function
End If
ReDim A(R(1, 1) To R(2, 1), R(1, 2) To (R(2, 2) - 1))
For i = R(1, 1) To R(2, 1)
    For j = R(1, 2) To (R(2, 2) - 1)
        A(i, j) = ay(i, j)
    Next j
Next i
de = mxdet(A)
If de = 0 Then
    debug_err "cramer", , "No solution for this case."
    Exit Function
End If
ReDim AA(R(1, 1) To R(2, 1), R(1, 2) To (R(2, 2) - 1))
ReDim x(R(1, 1) To R(2, 1))
For k = R(1, 1) To R(2, 1)
    For i = R(1, 1) To R(2, 1)
        For j = R(1, 2) To (R(2, 2) - 1)
            AA(i, j) = ay(i, j)
        Next j
        AA(i, k) = ay(i, R(2, 2))
    Next i
    Des = mxdet(AA)
    x(k) = Des / de
Next k
cramer = x
End Function

'******************************************************************************************************************************************************

'C3!: 3-Dimension array Functions
'Join n 2-Dimension arrays into a 3-dimension array (3rd,1st,2nd)
Public Function ayjoin2(ByVal ays As Variant) As Variant
Dim b As Boolean
Dim i As Long, j As Long, k As Long
Dim d As Integer, DA As Integer
Dim low As Long, up As Long
Dim A As Variant
Dim R As Variant, RA As Variant, RM As Variant
d = aydim(ays)
If d = 1 Then
    b = isay(ays, True)
    If b = True Then
        DA = aydims_m(ays)
        If DA = 2 Then
            RA = ayrges(ays)
            RM = ayrges_b(ays)
            ReDim R(LBound(RA, 1) To UBound(RA, 1), RM(1, 1) To RM(2, 1), RM(1, 2) To RM(2, 2)) As Variant
            For k = LBound(RA, 1) To UBound(RA, 1)
                For i = RA(k, 1, 1) To RA(k, 2, 1)
                    For j = RA(k, 1, 2) To RA(k, 2, 2)
                        R(k, i, j) = ays(k)(i, j)
                    Next j
                Next i
            Next k
            ayjoin2 = R
        ElseIf DA <> 1 Then
            debug_err "ayjoin1", , "At least one of the arrays in the jagged array is not one-dimensional."
        End If
    ElseIf b = False Then
        debug_err "ayjoin1", "njay"
    End If
ElseIf d <> 1 Then
    debug_err "ayjoin1", "1d"
End If
End Function

'******************************************************************************************************************************************************
'C4!: Jagged array Functions

'Dimensions of n arrays in a jagged array (limited to 3 dimensional array)
Public Function aydims(ByVal ay As Variant) As Variant
Dim i As Long, j As Long, k As Long
Dim b As Boolean
Dim d As Integer, d1 As Integer
Dim R As Variant
Dim ds As Variant
d = aydim(ay)
If d >= 1 Then
    b = isay(ay, True)
    If b = False Then
        aydims = d
    ElseIf b = True Then
        R = ayrge(ay)
        If d = 1 Then
            ReDim ds(R(1, 1) To R(2, 1))
            For i = R(1, 1) To R(2, 1)
                ds(i) = aydim(ay(i))
            Next i
        ElseIf d = 2 Then
            ReDim ds(R(1, 1) To R(2, 1), R(1, 2) To R(2, 2))
            For i = R(1, 1) To R(2, 1)
                For j = R(1, 2) To R(2, 2)
                    ds(i, j) = aydim(ay(i, j))
                Next j
            Next i
        ElseIf d = 3 Then
            ReDim ds(R(1, 1) To R(2, 1), R(1, 2) To R(2, 2), R(1, 3) To R(2, 3))
            For i = R(1, 1) To R(2, 1)
                For j = R(1, 2) To R(2, 2)
                    For k = R(1, 3) To R(2, 3)
                        ds(i, j, k) = aydim(ay(i, j, k))
                    Next k
                Next j
            Next i
        End If
    End If
    aydims = ds
ElseIf d = 0 Then
    aydims = d
End If
End Function

'Maximum Dimensions of n arrays in an jagged array
Public Function aydims_m(ByVal ay As Variant) As Variant
aydims_m = aytmax(aydims(ay))
End Function

'Ranges of a 1 dimension of a jagged array (Returns a 3-dimension array: m th array, Lbound vs Ubound, n th dimension)
Public Function ayrges(ByVal ays As Variant) As Variant
Dim i As Long, j As Long, k As Long
Dim R As Variant, RA As Variant, RB As Variant
Dim ds As Variant
Dim dsm As Long
If aydim(ays) = 1 Then
    If isay(ays, True) = True Then
        ds = aydims(ays)
        dsm = aymax(ds)
        RA = ayrge(ays)
        ReDim R(RA(1, 1) To RA(2, 1), 1 To 2, 1 To dsm)
        For i = RA(1, 1) To RA(2, 1)
            If isay(ays(i)) = True Then
                RB = ayrge(ays(i))
                For j = 1 To 2
                    For k = 1 To dsm
                        R(i, j, k) = RB(j, k)
                    Next k
                Next j
            End If
        Next i
        ayrges = R
    Else
        debug_err "ayrges", "njay"
    End If
Else
    debug_err "ayrges", "1d"
End If
End Function

'lowest Lbound and greatest Ubound in nth dimension of mth arrays in 1 dimension of a jagged array (returns a 2-dimension array: LLbound vs UUBound, nth dimension)
Public Function ayrges_b(ByVal ays As Variant) As Variant
Dim i As Long, j As Long
Dim low As Long, up As Long
Dim R As Variant, RA As Variant
RA = ayrges(ays)
low = LBound(RA, 3)
up = UBound(RA, 3)
ReDim R(1 To 2, low To up)
For j = low To up
    R(1, j) = aymin(ayredim(RA, Array(1, 1, j), 1))
    R(2, j) = aymax(ayredim(RA, Array(1, 2, j), 1))
Next j
ayrges_b = R
End Function

'Turn all values of specified position among the arrays in the jagged array (For 1 - dimensional jagged array only, up to 3 -dimensional arrays in the jagged array only)
Public Function jagtr1(ByVal ays As Variant, ByVal pos As Variant) As Variant
Dim i As Long
Dim b As Boolean
Dim d As Integer
Dim R As Variant, RP As Variant, RR As Variant
Dim S As Variant
b = isay(ays, True)
If b = True Then
    d = aydim(ays)
    If d = 1 Then
        R = ayrges(ays)
        RR = ayrge(R)
        ReDim S(RR(1, 1) To RR(2, 1)) As Variant
        d = aydim(pos)
        If d = 1 Then
            RP = ayrge(pos)
            If RP(2, 1) - RP(1, 1) = RR(2, 3) - RR(1, 3) Then
                d = RR(2, 3) - RR(1, 3) + 1
                If d = 1 Then
                    For i = RR(1, 1) To RR(2, 1)
                        S(i) = ays(i)(pos(RP(1, 1)))
                    Next i
                ElseIf d = 2 Then
                    For i = RR(1, 1) To RR(2, 1)
                        S(i) = ays(i)(pos(RP(1, 1)), pos(RP(1, 1) + 1))
                    Next i
                ElseIf d = 3 Then
                    For i = RR(1, 1) To RR(2, 1)
                        S(i) = ays(i)(pos(RP(1, 1)), pos(RP(1, 1) + 1), pos(RP(1, 1) + 2))
                    Next i
                End If
                jagtr1 = S
            Else
                debug_err "jaymax", , "Incorrect pos specified, please check."
            End If
        ElseIf d <> 1 Then
            debug_err "jaymax", , "pos must be a 1-dimensional array."
        End If
    Else
        debug_err "jaymax", "fna"
    End If
ElseIf b = False Then
    debug_err "jaymax", , "ays is not a jagged array."
End If
End Function

'******************************************************************************************************************************************************
'D!: Functions for objects

'Whether is directory or not
Public Function DirEx(ByVal path As String) As Boolean
DirEx = False
If Not dir(path, vbDirectory) = "" Then
    If GetAttr(path) And vbDirectory = vbDirectory Then
        DirEx = True
    End If
End If
End Function

'Whether a file exists
Public Function FileEx(path As String) As Boolean
On Error Resume Next
    FileEx = dir(path) <> ""
End Function

'Whether object exists
Public Function objex(ByRef obj As Object) As Boolean
Dim S As String
On Error Resume Next
S = obj.name
objex = (ERR.Number = 0)
ERR.Clear
End Function

'Get username
Public Function GetUserNameA() As String
  Dim lngResponse As Long
  Dim strUserName As String * 32
  lngResponse = GetUserName(strUserName, 32)
  GetUserNameA = Left(strUserName, InStr(strUserName, chr$(0)) - 1)
End Function

'Get computername
Public Function GetComputerNameA() As String
  Dim lngResponse As Long
  Dim strUserName As String * 32
  lngResponse = GetComputerName(strUserName, 32)
  GetComputerNameA = Left(strUserName, InStr(strUserName, chr$(0)) - 1)
End Function

'Find Window
Public Function FindWindowA(ByVal strClassName As String, ByVal strWindowName As String)
  Dim lngWnd As Long
  FindWindowA = FindWindow(strClassName, strWindowName)
End Function

'Returns exectuable for passed data file. (Note: strDataFile is full path of the file)
Public Function FindExecutableA(ByVal strDataFile As String, ByVal strDir As String) As String
Dim lngApp As Long
Dim strApp As String
strApp = Space(260)
lngApp = FindExecutable(strDataFile, strDir, strApp)
If lngApp > 32 Then
    FindExecutableA = strApp
Else
    FindExecutableA = "No matching application."
End If
End Function

'Get Attributes of a window
Public Function winatt(ByVal hwnd As Long, Optional ByVal ren As Boolean) As Variant
Dim leng As Long
Dim cname As String
Dim tname As String
Dim winid As Long
Dim Winrect As rect
Dim A() As Variant
If ren = False Then
    ReDim A(1 To 4) As Variant
ElseIf ren = True Then
    ReDim A(1 To 8) As Variant
End If
leng = GetWindowTextLength(hwnd)
cname = Space$(255)
tname = Space$(leng)
Call GetClassName(hwnd, cname, 255)
Call GetWindowText(hwnd, tname, leng + 1)
winid = CLng(GetWindowLong(hwnd, GWL_ID))
cname = TrimNull(cname)
tname = TrimNull(tname)
A(1) = hwnd
A(2) = cname
A(3) = tname
A(4) = winid
If rect = True Then
    Call GetWindowRect(hwnd, Winrect)
    A(5) = Winrect.Left
    A(6) = Winrect.Right
    A(7) = Winrect.Top
    A(8) = Winrect.Bottom
End If
winatt = A
End Function

'Active (Child) Windows Count
Public Function winct(Optional ByVal all As Boolean, Optional ByVal Child As Boolean, Optional ByVal hwnd As Long) As Long
    ct = 0
    If Child = False Then
        If all = False Then
            Call EnumWindows(AddressOf Vwinct, &H0)
        ElseIf all = True Then
            Call EnumWindows(AddressOf Awinct, &H0)
        End If
    ElseIf Child = True Then
        If hwnd <> 0 Then
            If all = False Then
                Call EnumChildWindows(hwnd, AddressOf Vwinct, 1)
            ElseIf all = True Then
                Call EnumChildWindows(hwnd, AddressOf Awinct, 1)
            End If
        ElseIf hwnd = 0 Then
            debug_err "winct", , "Please specify the hwnd when you Specify Child = True"
        End If
    End If
    winct = ct
End Function
    Public Function Vwinct(ByVal hwnd As Long, ByVal param As Long) As Long
    If IsWindowVisible(hwnd) Then
        ct = ct + 1
    End If
    Vwinct = 1
    End Function
    Public Function Awinct(ByVal hwnd As Long, ByVal param As Long) As Long
        ct = ct + 1
    Awinct = 1
    End Function

'Enumerate all the active (Child) windows (all: visible only or all; ren: includes size of the windows; Child: find Child windows)
Public Function enumwin(Optional ByVal all As Boolean, Optional ByVal ren As Boolean, Optional ByVal Child As Boolean, Optional ByVal hwnd As Long) As Variant
Dim c As Long
If Child = False Then
    c = winct(all)
ElseIf Child = True Then
    If hwnd <> 0 Then
        c = winct(all, Child, hwnd)
        If c = 0 Then
            debug_err "enumwin", , "There is no child windows for this hwnd:" & hwnd & "!!"
            Exit Function
        End If
    ElseIf hwnd = 0 Then
        debug_err "enumwin", , "Please specify the hwnd when you Specify Child = True"
        Exit Function
    End If
End If
If ren <> True Then
    ren = False
End If
rect = ren
If rect = False Then
    ReDim Win(1 To c, 1 To 4) As Variant
ElseIf rect = True Then
    ReDim Win(1 To c, 1 To 8) As Variant
End If
i = 1
If Child = False Then
    If all = False Then
        Call EnumWindows(AddressOf VEnumwin, &H0)
    ElseIf all = True Then
        Call EnumWindows(AddressOf AEnumwin, &H0)
    End If
ElseIf Child = True Then
    If all = False Then
        Call EnumChildWindows(hwnd, AddressOf VEnumwin, 1)
    ElseIf all = True Then
        Call EnumChildWindows(hwnd, AddressOf AEnumwin, 1)
    End If
End If
enumwin = Win
End Function
    Public Function VEnumwin(ByVal hwnd As Long, ByVal param As Long) As Long
    Dim leng As Long
    Dim cname As String
    Dim tname As String
    Dim winid As Long
    Dim Winrect As rect
    If IsWindowVisible(hwnd) Then
        leng = GetWindowTextLength(hwnd)
        cname = Space$(255)
        tname = Space$(leng)
        Call GetClassName(hwnd, cname, 255)
        Call GetWindowText(hwnd, tname, leng + 1)
        winid = CLng(GetWindowLong(hwnd, GWL_ID))
        cname = TrimNull(cname)
        tname = TrimNull(tname)
        Win(i, 1) = hwnd
        Win(i, 2) = cname
        Win(i, 3) = tname
        Win(i, 4) = winid
        If rect = True Then
            Call GetWindowRect(hwnd, Winrect)
            Win(i, 5) = Winrect.Left
            Win(i, 6) = Winrect.Right
            Win(i, 7) = Winrect.Top
            Win(i, 8) = Winrect.Bottom
        End If
        i = i + 1
    End If
    VEnumwin = 1
    End Function
    Public Function AEnumwin(ByVal hwnd As Long, ByVal param As Long) As Long
    Dim leng As Long
    Dim cname As String
    Dim tname As String
    Dim winid As Long
    Dim Winrect As rect
    leng = GetWindowTextLength(hwnd)
    cname = Space$(255)
    tname = Space$(leng)
    Call GetClassName(hwnd, cname, 255)
    Call GetWindowText(hwnd, tname, leng + 1)
    winid = CLng(GetWindowLong(hwnd, GWL_ID))
    cname = TrimNull(cname)
    tname = TrimNull(tname)
    Win(i, 1) = hwnd
    Win(i, 2) = cname
    Win(i, 3) = tname
    Win(i, 4) = winid
    If rect = True Then
        Call GetWindowRect(hwnd, Winrect)
        Win(i, 5) = Winrect.Left
        Win(i, 6) = Winrect.Right
        Win(i, 7) = Winrect.Top
        Win(i, 8) = Winrect.Bottom
    End If
    i = i + 1
    AEnumwin = 1
    End Function

'Return system's temporary folder.
Public Function GetTempPathA() As String
Dim strPath As String * 512
Dim lgnPath As Long
lgnPath = GetTempPath(512, strPath)
GetTempPathA = Left(strPath, InStr(1, strPath, vbNullChar))
End Function

'Whether a file exists
Public Function FileExist(ByVal localpath As String) As Boolean
Dim fn As String
Dim x As String
x = dir(localpath)
fn = Right(localpath, (Len(localpath) - InStrRev(localpath, "\")))
If x <> "" Then
    If x = fn Then
        FileExist = True
    ElseIf x <> fn Then
        debug_err "Fileex", , "localpath is not the path of a file, please check!"
        Exit Function
    End If
ElseIf x = "" Then
    FileExist = False
End If
End Function

'Whether a folder exists
Public Function Fdrex(ByVal fdr As String) As Boolean
Dim x As String
Dim fn As String
Dim path As String
If Right(fdr, 1) = "\" Then
    path = fdr
ElseIf Right(fdr, 1) <> "\" Then
    If dir(fdr, vbDirectory) = "" Then
        path = fdr & "\"
    ElseIf dir(fdr, vbDirectory) <> "" Then
        path = fdr
    End If
End If
x = dir(path, vbDirectory)
fn = Right(path, (Len(path) - InStrRev(path, "\")))
If x <> "" Then
    If x <> fn Then
        Fdrex = True
    ElseIf x = fn Then
        debug_err "Fileex", , "localpath is not the path of a folder, please check!"
        Exit Function
    End If
Else
End If
End Function

'Count file in a path
Public Function file_ct(ByVal path As String) As Long
Dim FSO As New FileSystemObject
Dim FIL As File
Dim fdr As folder
Dim k As Long
Set fdr = FSO.Getfolder(path)
k = 0
For Each FIL In fdr.Files
    k = k + 1
Next FIL
file_ct = k
End Function

'Get file properties
Public Function Get_File_prop(ByVal path As String) As Variant
Dim z(1 To 2) As Long
Dim FSO As New FileSystemObject
Dim F(1 To 7) As Variant
Dim FIL As File

Set FIL = FSO.GetFile(path)
F(1) = FIL.name
z(1) = Len(F(1))
z(2) = InStrRev(F(1), ".")
If z(2) <> 0 Then
    F(2) = Left(F(1), (z(2) - 1))
    F(3) = Right(F(1), z(1) - z(2))
    F(4) = FIL.size
    F(5) = FIL.DateCreated
    F(6) = FIL.DateLastModified
    F(7) = FIL.DateLastAccessed
End If
Get_File_prop = F
End Function

'Get file properties in a path
'(1: Full File name, 2: File name, 3: File name Extension, 4.Size, 5. Date Created, 6. File last modified, 7. Date last Accessed)
Public Function Get_files(ByVal path As String) As Variant
Dim FSO As New FileSystemObject
Dim F() As Variant
Dim FIL As File
Dim fdr As folder
Dim k As Long
Dim fn As Long
Dim z(1 To 3) As Long
If Right(path, 1) <> "\" Then
    path = path & "\"
ElseIf Right(path, 1) = "\" Then
    path = path
End If
Set fdr = FSO.Getfolder(path)
k = 1
For Each FIL In fdr.Files
    k = k + 1
Next FIL
fn = k - 1
ReDim F(1 To fn, 1 To 7)
k = 1
For Each FIL In fdr.Files
    F(k, 1) = FIL.name
    z(1) = Len(F(k, 1))
    z(2) = InStrRev(F(k, 1), ".")
    If z(2) <> 0 Then
        F(k, 2) = Left(F(k, 1), (z(2) - 1))
        F(k, 3) = Right(F(k, 1), z(1) - z(2))
        F(k, 4) = FIL.size
        F(k, 5) = FIL.DateCreated
        F(k, 6) = FIL.DateLastModified
        F(k, 7) = FIL.DateLastAccessed
    End If
    k = k + 1
Next FIL
Get_files = F
End Function

'Get Folder list in a path
'(1: Name, 2: Path, 3:Size, 4. Date Created, 5. File last modified, 6. Date last Accessed, 7.Folder Count of the folder, 8. File Count of the folder)
Public Function get_fdr(ByVal path As String) As Variant
Dim k As Long
Dim FSO As New FileSystemObject
Dim F() As Variant
Dim fdr As folder, fdr1 As folder
Dim fn As Long
If Right(path, 1) <> "\" Then
    path = path & "\"
ElseIf Right(path, 1) = "\" Then
    path = path
End If
Set fdr = FSO.Getfolder(path)
k = 1
For Each fdr1 In fdr.SubFolders
    k = k + 1
Next fdr1
fn = k - 1
ReDim F(1 To fn, 1 To 8)
k = 1
For Each fdr1 In fdr.SubFolders
    On Error Resume Next
    F(k, 7) = fdr1.SubFolders.Count
    F(k, 8) = fdr1.Files.Count
    F(k, 1) = fdr1.name
    F(k, 2) = fdr1.path
    If ERR.Number = 0 Then
        F(k, 3) = fdr1.size
    ElseIf ERR.Number > 0 Then
        F(k, 3) = "N/A"
    End If
    ERR.Clear
    On Error GoTo 0
    F(k, 4) = fdr1.DateCreated
    F(k, 5) = fdr1.DateLastModified
    F(k, 6) = fdr1.DateLastAccessed
    k = k + 1
Next fdr1
get_fdr = F
End Function

Function Getfolder(ByVal path As String) As Boolean
Dim FSO As New FileSystemObject
Dim fdr As folder, fdr1 As folder
Set fdr = FSO.Getfolder(path)

For Each fdr1 In fdr.SubFolders
    On Error Resume Next
    Debug.Print fdr1.path
    Getfolder fdr1.path
Next fdr1
End Function

'Copy File
Function copyfile(ByVal fp As String, ByVal Des As String, Optional ByVal ovrwrt As Boolean, Optional Debug_msg As String) As Boolean
Dim exf As Boolean, exd As Boolean
exf = FileEx(fp)
If exf = True Then
    exd = FileEx(Des)
    If exd = False Then
    Else
        If ovrwrt = False Then
            Debug_msg = debug_err("copyfile", , "Destination file exists : " & Des)
            Exit Function
        Else
        End If
    End If
    On Error Resume Next
    FileCopy fp, Des
    If ERR.Number = 70 Then
        Debug_msg = debug_err("copyfile", , "There is no permission to write the file to " & Des)
    End If
Else
    Debug_msg = debug_err("copyfile", , "File to be copyed does not exist : " & fp)
    Exit Function
End If
copyfile = True
End Function

'Move file
Function movefile(ByVal fp As String, ByVal fdr As String, Optional ByVal ovrwrt As Boolean) As Boolean
Dim exf As Boolean, exfr As Boolean, exd As Boolean
Dim path As String, Des As String, fn As String
exf = FileEx(fp)
If exf = True Then
    exfr = Fdrex(fdr)
    If exfr = True Then
        If Right(fdr, 1) = "\" Then
            path = fdr
        ElseIf Right(fdr, 1) <> "\" Then
            path = fdr & "\"
        End If
        fn = Right(fp, (Len(fp) - InStrRev(fp, "\")))
        Des = path & fn
        exd = FileEx(Des)
        If exd = False Then
            Name fp As Des
        Else
            If ovrwrt = False Then
                debug_err "movefile", , "Destination file exists"
                Exit Function
            ElseIf ovrwrt = True Then
                Name fp As Des
            End If
        End If
    Else
        debug_err "movefile", , "The folder does not exist"
        Exit Function
    End If
Else
    debug_err "movefile", , "File to be moved does not exist"
    Exit Function
End If
movefile = True
End Function

'Rename File (include sub file name)
Public Function rename(ByVal old_name As String, ByVal new_name As String, ByVal folder As String)
If Right(folder, 1) <> "\" Then
    folder = folder & "\"
End If
Name folder & old_name As folder & new_name
End Function

'Set Create/Modified/Access date & Time of a file
Public Function SetFileDateTime(ByVal filename As String, ByVal Dat As String, ByVal datetime_Typ As String) As Boolean
Dim lFileHnd As Long
Dim lRet As Long
Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

If dir(filename) = "" Then Exit Function
If Not IsDate(Dat) Then Exit Function

With typSystemTime
    .wYear = year(Dat)
    .wMonth = month(Dat)
    .wDay = day(Dat)
    .wDayOfWeek = Weekday(Dat) - 1
    .wHour = Hour(Dat)
    .wMinute = Minute(Dat)
    .wSecond = Second(Dat)
End With

lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)
lFileHnd = CreateFile(filename, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
'lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, typFileTime)
If datetime_Typ = "C" Then
    lRet = SetFileTimeCreate(lFileHnd, typFileTime, ByVal 0&, ByVal 0&)
ElseIf datetime_Typ = "M" Then
    lRet = SetFileTimeLastModified(lFileHnd, ByVal 0&, ByVal 0&, typFileTime)
ElseIf datetime_Typ = "A" Then
    lRet = SetFileTimeLastAccess(lFileHnd, ByVal 0&, typFileTime, ByVal 0&)
End If

CloseHandle lFileHnd
SetFileDateTime = lRet > 0
End Function

'Zip a file/path to a file
Function CreateZipFile(path As Variant, ZipFileName As Variant) As Boolean
Dim ShellApp As Object

Open ZipFileName For Output As #1
Print #1, chr$(80) & chr$(75) & chr$(5) & chr$(6) & String(18, 0)
Close #1

Set ShellApp = CreateObject("Shell.Application")
ShellApp.Namespace(ZipFileName).CopyHere ShellApp.Namespace(path).Items
End Function

'upzip a file to a folder
Public Function unzip(ByVal zip_path As Variant, ByVal target_folder As Variant, ByVal overwrite As Boolean) As Boolean
Dim FSO As Object
Dim oApp As Object
Dim opt As Variant
If Right(target_folder, 1) <> "\" Then
    target_folder = target_folder & "\"
End If
If overwrite = False Then
    opt = 0
Else
    opt = 20
End If
Set oApp = CreateObject("Shell.Application")
oApp.Namespace(target_folder).CopyHere oApp.Namespace(zip_path).Items, opt
End Function

'Get Txt file to String (csv file accpeted too)
Public Function txtstr(ByVal FilePath As String) As String
Open FilePath For Binary As #1
txtstr = Space$(LOF(1))
Get #1, , txtstr
Close #1
End Function

'Save String to file
Public Function strtxt(ByVal str As String, ByVal path As String, ByVal full_name As String, Optional ByVal unicode As Boolean) As Boolean
Dim obj As Object
Dim txtstream As Variant
Set obj = CreateObject("Scripting.FileSystemObject")
If Right(path, 1) <> "\" Then
    path = path & "\"
End If
If unicode = False Then
    Set txtstream = obj.OpenTextFile(path & full_name, Create:=True, IOMode:=ForAppending)
Else
    Set txtstream = obj.OpenTextFile(path & full_name, Create:=True, IOMode:=ForAppending, format:=TristateTrue)
End If
txtstream.WriteLine (str)
txtstream.Close
End Function

'csv file to 2-dimensional array
Public Function csvay(ByVal FilePath As String) As Variant
Dim i As Long, j As Long
Dim str As String, str1 As Variant, str2 As Variant
Dim A As Variant
Dim R As Variant, c As Variant, CC As Variant
str = txtstr(FilePath)
str1 = wDay(str, (chr(13) & chr(10)))
R = ayrge(str1)
ReDim str2(R(1, 1) To R(2, 1))
For i = R(1, 1) To R(2, 1)
    str2(i) = wDay(TrimG(str1(i), chr(10)), ",")
Next i
c = ayrges_b(str2)
ReDim A(R(1, 1) To R(2, 1), c(1, 1) To c(2, 1))
For i = R(1, 1) To R(2, 1)
    CC = ayrge(str2(i))
    For j = c(1, 1) To c(2, 1)
        If j <= CC(2, 1) Then
            A(i, j) = str2(i)(j)
        End If
    Next j
Next i
csvay = A
End Function

'csv file to 2-dimensional array (Not call any function)
Public Function csvay_f(ByVal FilePath As String, Optional ByVal break As String) As Variant
Dim i As Long, j As Long, S As Long, t As Long, S2 As Long, t2 As Long
Dim i_ct As Long, j_ct As Long
Dim str As String, str1 As Variant, str2 As Variant
Dim A As Variant
Dim NL As String
NL = (chr(10))
If break = "" Then break = ","
'Get string from the path of the file
Open FilePath For Binary As #1
str = Space$(LOF(1))
Get #1, , str
Close #1
'Count the array size
i = 1
S = 1
t = 1
i_ct = 1
j_ct = 1
Do While t > 0
    t = InStr(i, str, NL)
    j = 0
    If t > 0 Then
        i = S
        Do
            i = InStr(i, str, break)
            If i > 0 Then
                j = j + 1
                i = i + 1
            End If
        Loop While i < t
        If j > j_ct Then
            j_ct = j
        End If
        i_ct = i_ct + 1
    Else
        Do While i > 0
            i = InStr(i, str, break)
            If i > 0 Then
                j = j + 1
                i = i + 1
            End If
        Loop
        If j > j_ct Then
            j_ct = j
        End If
    End If
    S = t + 1
Loop

ReDim A(1 To i_ct, 1 To j_ct) As Variant
'Split the str into an array
i = 1
S = 1
t = 1
i_ct = 1
S2 = 0
Do While t > 0
    t = InStr(i, str, NL)
    j = 0
    If t > 0 Then
        i = S
        Do
            i = InStr(i, str, break)
            t2 = i
            If t2 > t Then
                t2 = t
            End If
            If t2 > 0 Then
                j = j + 1
                A(i_ct, j) = mid(str, S2 + 1, t2 - S2 - 1)
                i = i + 1
                S2 = t2
            End If
        Loop While t2 < t
        If j > j_ct Then
            j_ct = j
        End If
        i_ct = i_ct + 1
    ElseIf t = 0 Then
        j = j + 1
        A(i_ct, j) = mid(str, S + 1, i - S - 2)
        S2 = i - 1
        t2 = S2
        Do While t2 > 0
            i = InStr(i, str, break)
            t2 = i
            If t2 > 0 Then
                j = j + 1
                A(i_ct, j) = mid(str, S2 + 1, t2 - S2 - 1)
                i = i + 1
                S2 = t2
            ElseIf t2 = 0 Then
                j = j + 1
                A(i_ct, j) = mid(str, S2 + 1, Len(str) - S2)
            End If
        Loop
    End If
    S = t + 1
    S2 = S
Loop
csvay_f = A
End Function

Function Gen_dummy_file(ByVal path As String, ByVal size As Long) As Boolean
Dim i As Long, n As Long
Dim str As String
Open path For Binary As #1
For i = 1 To size
n = Int(Rnd() * 26) + 65
str = chr(n)
Put #1, , str
Next i
Close #1
End Function

'D1! Office Application
'Whether the specify application exist
Public Function appex(ByVal App As Application) As Boolean
Dim str As String
On Error Resume Next
str = App.name
If ERR.Number = 0 Then
    appex = True
Else
    appex = False
End If
End Function

'Adding reference to a MS office Application
Function addreference(ByVal wb As Object, ByVal path As String) As Boolean
Dim vbProject As Object
Set vbProject = wb.vbProject
On Error Resume Next
vbProject.References.AddFromFile path
If ERR.Number > 0 Then
    If ERR.Nunber = 32813 Then
        debug_err "addreference", , "The reference has been already exists."
        On Error GoTo 0
        Exit Function
    End If
    debug_err "addreference", , "Unknown Error!"
    Exit Function
End If
On Error GoTo 0
addreference = True
End Function

'Deleting reference of a MS office Application
Function delreference(ByVal wb As Object, ByVal name As String) As Boolean
Dim ref As Object
Dim vbProj As Object
Set vbProj = wb.vbProject
On Error Resume Next
Set ref = vbProj.References(name)
If ERR.Number > 0 Then
    If ERR.Number = 9 Then
        debug_err "delreference", , "The reference has not been found."
        On Error GoTo 0
        Exit Function
    End If
    debug_err "delreference", , "Unknown Error!"
    Exit Function
End If
vbProj.References.Remove ref
On Error GoTo 0
delreference = True
End Function

'Current Excel Application
Public Function Thisapp() As Application
Set Thisapp = ThisWorkbook.Application
End Function

'Create Excel Application with Settings
Public Function xlappcreate(Optional ByVal xlvisible As Boolean, Optional xldisplay As Boolean) As Excel.Application
Dim xlapp As Application
Set xlapp = CreateObject("Excel.Application")
xlapp.Visible = xlvisible
xlapp.Application.DisplayAlerts = xldisplay
Set xlappcreate = xlapp
End Function

'Whether any excel application exist
Public Function xlappex() As Boolean
Dim ct As Integer
ct = xlappct()
If ct > 0 Then
    xlappex = True
ElseIf ct = 0 Then
    xlappex = False
End If
End Function

'Count for excel applications
Public Function xlappct() As Long
Dim GUID&(0 To 3), acc As Object, hwnd, hwnd2, hwnd3
Dim ct As Long
GUID(0) = &H20400
GUID(1) = &H0
GUID(2) = &HC0
GUID(3) = &H46000000
Do
    hwnd = FindWindowExA(0, hwnd, "XLMAIN", vbNullString)
    If hwnd = 0 Then Exit Do
        hwnd2 = FindWindowExA(hwnd, 0, "XLDESK", vbNullString)
        hwnd3 = FindWindowExA(hwnd2, 0, "EXCEL7", vbNullString)
    If AccessibleObjectFromWindow(hwnd3, &HFFFFFFF0, GUID(0), acc) = 0 Then
        ct = ct + 1
    End If
Loop
xlappct = ct
End Function

'Get all excel application
Public Function Getxlapps() As Variant
Dim i As Long
Dim GUID&(0 To 3), acc As Object, hwnd, hwnd2, hwnd3
Dim A As Variant
Dim ct As Long
GUID(0) = &H20400
GUID(1) = &H0
GUID(2) = &HC0
GUID(3) = &H46000000
ct = xlappct()
If ct > 0 Then
    ReDim A(1 To ct) As Variant
    i = 1
    Do
        hwnd = FindWindowExA(0, hwnd, "XLMAIN", vbNullString)
        If hwnd = 0 Then Exit Do
            hwnd2 = FindWindowExA(hwnd, 0, "XLDESK", vbNullString)
            hwnd3 = FindWindowExA(hwnd2, 0, "EXCEL7", vbNullString)
        If AccessibleObjectFromWindow(hwnd3, &HFFFFFFF0, GUID(0), acc) = 0 Then
            Set A(i) = acc.Application
            i = i + 1
        End If
    Loop
    Getxlapps = A
ElseIf ct = 0 Then
    debug_err "Getxlapps", , "No Excel application is active"
End If
End Function

'Close all invisible excel application (return the number of xlapps closed)(Failed)
Public Function closexlapps() As Long
Dim i As Long
Dim ct As Long
Dim xlapps As Variant, xlapp As Variant, txlapp As Application
Set txlapp = ThisWorkbook.Application
xlapps = Getxlapps()
ct = 0
For i = LBound(xlapps, 1) To UBound(xlapps, 1)
    If xlapps(i).Visible = False Then
        xlapps(i).DisplayAlerts = False
        xlapps(i).Quit
        ct = ct + 1
    End If
Next i
closexlapps = ct
End Function

'Whether access application exist
Public Function accappex() As Boolean
Dim ct As Integer
ct = accappct()
If ct > 0 Then
    accappex = True
ElseIf ct = 0 Then
    accappex = False
End If
End Function

'Count for Access applications
Public Function accappct() As Long
Dim GUID&(0 To 3), acc As Object, hwnd
Dim ct As Long
GUID(0) = &H20400
GUID(1) = &H0
GUID(2) = &HC0
GUID(3) = &H46000000
ct = 0
Do
    hwnd = FindWindowExA(0, hwnd, "OMAIN", vbNullString)
    If hwnd = 0 Then Exit Do
    If AccessibleObjectFromWindow(hwnd, &HFFFFFFF0, GUID(0), acc) = 0 Then
        ct = ct + 1
    End If
Loop
accappct = ct
End Function

'Get all access application
Public Function Getaccapps() As Variant
Dim i As Long
Dim GUID&(0 To 3), acc As Object, hwnd
Dim A As Variant
Dim A2() As Access.Application
Dim App As Access.Application
Dim ct As Long
GUID(0) = &H20400
GUID(1) = &H0
GUID(2) = &HC0
GUID(3) = &H46000000
ct = accappct()
If ct > 0 Then
    ReDim A(1 To ct) As Variant
    i = 1
    Do
        hwnd = FindWindowExA(0, hwnd, "OMAIN", vbNullString)
        If hwnd = 0 Then Exit Do
        If AccessibleObjectFromWindow(hwnd, &HFFFFFFF0, GUID(0), acc) = 0 Then
            Set A(i) = acc.Application
            i = i + 1
        End If
    Loop
    Getaccapps = A
ElseIf ct = 0 Then
    debug_err "Getaccapps", , "No Access application is active"
End If
End Function

'Count for Word applications
Public Function wordappct() As Long
Dim GUID&(0 To 3), acc As Object, hwnd, hwnd2, hwnd3, hwnd4
Dim ct As Long
GUID(0) = &H20400
GUID(1) = &H0
GUID(2) = &HC0
GUID(3) = &H46000000
ct = 0
Do
    hwnd = FindWindowExA(0, hwnd, "OpusApp", vbNullString)
    If hwnd = 0 Then Exit Do
    hwnd2 = FindWindowExA(hwnd, 0, "_wwF", vbNullString)
    hwnd3 = FindWindowExA(hwnd2, 0, "_wwB", vbNullString)
    hwnd4 = FindWindowExA(hwnd3, 0, "_wwG", vbNullString)
    If AccessibleObjectFromWindow(hwnd4, &HFFFFFFF0, GUID(0), acc) = 0 Then
        ct = ct + 1
    End If
Loop
wordappct = ct
End Function

'Get all word application as a Collection
Public Function Getwordapps() As Variant
Dim GUID&(0 To 3), acc As Object, hwnd, hwnd2, hwnd3, hwnd4
Dim A As Collection
Dim A2() As Word.Application
Dim App As Word.Application
Dim ct As Long
GUID(0) = &H20400
GUID(1) = &H0
GUID(2) = &HC0
GUID(3) = &H46000000
 
Set A = New Collection
Do
    hwnd = FindWindowExA(0, hwnd, "OpusApp", vbNullString)
    If hwnd = 0 Then Exit Do
    hwnd2 = FindWindowExA(hwnd, 0, "_wwF", vbNullString)
    hwnd3 = FindWindowExA(hwnd2, 0, "_wwB", vbNullString)
    hwnd4 = FindWindowExA(hwnd3, 0, "_wwG", vbNullString)
    If AccessibleObjectFromWindow(hwnd4, &HFFFFFFF0, GUID(0), acc) = 0 Then
        A.add acc.Application
    End If
Loop
ct = A.Count
If ct > 0 Then
    ReDim A2(1 To ct) As Word.Application
    i = 1
    For Each App In A
        Set A2(i) = App
        i = i + 1
    Next App
    Getwordapps = A2
ElseIf ct = 0 Then
    debug_err "Getwordapps", , "No Word application is active"
End If
End Function

'Count for IE applications
Public Function iect() As Long
Dim i As Long
Dim objShell As Object
Dim ow As Variant
Set objShell = CreateObject("Shell.Application")
i = 0
For Each ow In objShell.Windows
    If (InStr(1, ow, "Internet Explorer", vbTextCompare)) Then
        i = i + 1
    End If
Next
iect = i
End Function

'Get IE applications
Public Function getieapp() As Variant
Dim i As Long
Dim ct As Long
Dim objShell As Object
Dim ie As Variant, ies() As SHDocVw.InternetExplorer
Set objShell = CreateObject("Shell.Application")
ct = iect()
ReDim ies(1 To ct) As SHDocVw.InternetExplorer
i = 1
For Each ie In objShell.Windows
    If (InStr(1, ie, "Internet Explorer", vbTextCompare)) Then
        Set ies(i) = ie
        i = i + 1
    End If
Next
getieapp = ies
End Function

'Return loading time of ie
Public Function ieload(ByVal ieapp As Object) As Double
Dim t(1 To 3) As Double
t(1) = Timer()
Do While ieapp.Busy And Not (ieapp.readyState = READYSTATE_COMPLETE)
    DoEvents
Loop
t(2) = Timer()
t(3) = t(2) - t(1)
End Function

'Goto google searching
Public Function Googlesearch(ByVal str As String) As Boolean
Dim obj As Object
Dim ie As New SHDocVw.InternetExplorer
ie.Visible = True
ie.navigate "http://www.google.com"
Do While ie.Busy
    DoEvents
Loop
Do While ie.Busy And Not (ie.readyState = READYSTATE_COMPLETE)
    DoEvents
Loop
Set obj = ie.Document.getElementById("lst-ib")
obj.Value = str
Do While ie.Busy
    DoEvents
Loop
Do While ie.Busy And Not (ie.readyState = READYSTATE_COMPLETE)
    DoEvents
Loop
ie.Document.all("btnG").Click
End Function

'Open exe Program
Public Function openexe(ByVal path As String) As Long
Dim S As Long
On Error Resume Next
If dir(path) <> "" Then
    S = Shell(path, vbMinimizedFocus)
ElseIf dir(path) = "" Then
    debug_err "openexe", , "The path is invalid, please check."
    Exit Function
End If
openexe = S
End Function

'Open ie Application
Public Function openie(Optional ByVal hide As Boolean) As Boolean
Dim ct As Integer, ct1 As Integer
Dim ieapps As Variant, ie As Variant
Shell ("C:\Program Files\Internet Explorer\iexplore.exe")
ct1 = iect()
If hide = True Then
    Do Until ct = (ct1 + 1)
        DoEvents
        ct = iect()
    Loop
    ieapps = getieapp()
    For Each ie In ieapps
        ie.Visible = False
    Next ie
End If
End Function

'Close all ie Application
Public Function closeieall() As Boolean
Dim ieapps As Variant, ie As Variant
ieapps = getieapp()
For Each ie In ieapps
    ie.Quit
Next ie
End Function

'Open Chrome Application
Public Function openchrome(Optional ByVal Url As String) As Boolean
Dim Clink As String, Flink As String
Clink = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
If Url = vbNullString Then
    Flink = Clink
ElseIf Url <> vbNullString Then
    Flink = Clink & " url: " & Url
End If
Shell (Flink)
openchrome = True
End Function

'Save the web file into a specific path
Function SaveWebFile(ByVal webpath As String, ByVal localDir As String, Optional ByVal overwrite As Boolean, Optional filename As String) As Boolean
    Dim oXMLHTTP As Object
    Dim vFF As Long
    Dim oResp() As Byte
    Dim localpath As String
    Dim webfile As String
    Dim subfile As String
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    oXMLHTTP.Open "GET", webpath, False
    oXMLHTTP.Send
    Do While oXMLHTTP.readyState <> 4
        DoEvents
    Loop
    oResp = oXMLHTTP.ResponseBody
    subfile = Right(webpath, (Len(webpath) - InStrRev(webpath, ".")))
    If Right(localDir, 1) <> "\" Then localDir = localDir & "\"
    If filename = "" Then
        localpath = localDir & Right(webpath, (Len(webpath) - InStrRev(webpath, "/")))
    ElseIf filename <> "" Then
        localpath = localDir & filename & "." & subfile
    End If
    If overwrite = False Then
        If dir(localpath) <> "" Then
            debug_err "Savewebfile", , "The file exists, please check!"
        End If
    ElseIf overwrite = True Then
        If dir(localpath) <> "" Then kill localpath
    End If
    vFF = FreeFile
    Open localpath For Binary As #vFF
    Put #vFF, , oResp
    Close #vFF
    Set oXMLHTTP = Nothing
End Function

'Get the html of the specified webpath
Function gethtml(ByVal webpath As String) As String
Dim v As Variant
Dim XMLHTTP As Object
Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
XMLHTTP.Open "GET", webpath
XMLHTTP.Send
gethtml = XMLHTTP.responseText
End Function

'Get html file by html strings
Function htmldoc(ByVal html_s As String) As HTMLDocument
Dim h As New HTMLDocument
h.Body.innerHTML = html_s
Set htmldoc = h
End Function

'Remove sysMenu of a form
Public Function RemoveCloseButton(frm As Object) As Boolean
Dim lngStyle As Long
Dim lngHWnd As Long
    lngHWnd = FindWindow(vbNullString, frm.Caption)
    lngStyle = GetWindowLong(lngHWnd, GWL_STYLE)
    If lngStyle And WS_SYSMENU > 0 Then
        SetWindowLong lngHWnd, GWL_STYLE, (lngStyle And Not WS_SYSMENU)
    End If
End Function

'Form button (Enable/ Disable min or max button, Disable close button only)
Public Function frmbutton(ByVal frm As Object, ByVal button As String, Optional ByVal able As Boolean) As Boolean
Dim hwnd As Long
Dim WinStyle As Long
Dim Result As Long
hwnd = FindWindow(vbNullString, frm.Caption)
WinStyle = GetWindowLong(hwnd, GWL_STYLE)
hmenu = GetSystemMenu(hwnd, 0)
If button = "MIN" Then
    If able = True Then
        If (WinStyle And MIN_BOX) = 0 Then
            SetWindowLong hwnd, GWL_STYLE, (WinStyle Or MIN_BOX)
        End If
    ElseIf able = False Then
        If (WinStyle And MIN_BOX) = MIN_BOX Then
            SetWindowLong hwnd, GWL_STYLE, (WinStyle And Not (MIN_BOX))
        End If
    End If
ElseIf button = "MAX" Then
    If able = True Then
        If (WinStyle And MAX_BOX) = 0 Then
            SetWindowLong hwnd, GWL_STYLE, (WinStyle Or MAX_BOX)
        End If
    ElseIf able = False Then
        If (WinStyle And MAX_BOX) = MAX_BOX Then
            SetWindowLong hwnd, GWL_STYLE, (WinStyle And Not (MAX_BOX))
        End If
    End If
ElseIf button = "CLOSE" Then
    If able = False Then
        RemoveMenu hmenu, SC_CLOSE, 0&
    ElseIf able = True Then
        RemoveMenu hmenu, SC_CLOSE, 0&
    End If
End If
DrawMenuBar hwnd
End Function


'******************************************************************************************************************************************************
'E!: functions for excel

'Whether workbook exists
Public Function wbex(ByVal wb As Workbook) As Boolean
Dim str As String
On Error Resume Next
str = wb.name
If ERR = 0 Then
    wbex = True
ElseIf ERR > 0 Then
    wbex = False
End If
On Error GoTo 0
End Function

'Whether any workbook exists or not in the specify Application (search by name)
Public Function wbexn(ByVal wb_name As String, ByRef xlapp As Application) As Boolean
Dim wb As Workbook
For Each wb In xlapp.Workbooks
    If (wb.name = wb_name) = True Then
        wbexn = True
        Exit Function
    End If
Next
End Function

'Count for active workbooks
Public Function wbct(Optional ByVal App As Application, Optional byapp As Boolean) As Integer
Dim i As Long
Dim appa As Variant, apps As Variant
Dim wb As Workbook
Dim ex As Boolean
Dim c As Variant, ca As Variant
Dim ct As Long
ex = appex(App)
If ex = True Then
    For Each wb In App.Workbooks
        c = c + 1
    Next wb
Else
    ct = xlappct
    If ct = 1 Then
        apps = Getxlapps
        c = wbct(apps(1))
    ElseIf ct > 1 Then
        apps = Getxlapps()
        ReDim ca(1 To ct) As Variant
        For Each appa In apps
            i = i + 1
            ca(i) = ca(i) + 1
        Next appa
        If byapp = True Then
            c = ca
        Else
            c = aysum(ca)
        End If
    End If
End If
wbct = c
End Function

'Get all active workbooks
Public Function getwbs(Optional ByVal App As Application, Optional byapp As Boolean) As Variant
Dim i As Long, j As Long
Dim ex As Boolean
Dim ct As Variant
Dim apps As Variant, appa As Variant
Dim wb As Workbook, A() As Workbook
ex = appex(App)
If ex = True Then
    ct = wbct(App)
    ReDim A(1 To ct) As Workbook
    For Each wb In App.Workbooks
        i = i + 1
        Set A(i) = wb
    Next wb
    getwbs = A
Else
    apps = Getxlapps()
    If byapp = False Then
        ct = wbct(App)
        ReDim A(1 To ct) As Workbook
        For Each appa In apps
            For Each wb In appa.Workbooks
                i = i + 1
                Set A(i) = wb
            Next wb
        Next appa
    Else
        ct = wbct(App, True)
        ReDim A(1 To UBound(ct, 1), 1 To UBound(ct, 2))
        For Each appa In apps
            i = i + 1
            For Each wb In appa.Workbooks
                j = j + 1
                Set A(i, j) = wb
            Next wb
        Next appa
    End If
    getwbs = A
End If
End Function

'Set workbook of an application
Public Function wbset(ByVal wb_name As String, ByVal xlapp As Application, Optional ByVal con As Boolean) As Workbook
Dim wb As Workbook
Dim b As Boolean
For Each wb In xlapp.Workbooks
    If con = False Then
        b = (wb.name = wb_name)
    Else
        b = (InStr(1, wb.name, wb_name) <> 0)
    End If
    If b = True Then
        Set wbset = wb
        Exit Function
    End If
Next wb
End Function

'Save Workbook
Public Function wbsave(ByVal wb As Workbook, ByVal fdr_path As String, ByVal name As String, Optional ByVal format As Variant, Optional ByVal Pw As String, Optional ByVal wrpw As String, Optional ByVal RdOnly As Boolean) As Workbook
Dim wb1 As Workbook
Dim xlapp As Application
On Error Resume Next
Application.DisplayAlerts = False
save:
wb.SaveAs fdr_path & "\" & name, format, Pw
If ERR.Number = 1004 Then
    Set xlapp = wb.Application
    xlapp.Application.DisplayAlerts = False
    Set wb1 = wbset(name, xlapp, True)
    wb1.Close True
    On Error GoTo 0
    GoTo save
End If
End Function

'change one excel file type
Public Function xlstyp(ByVal input_file As String, ByVal typ_b As String, ByVal typ_a As String, Optional ByVal output_file As String, Optional ByVal del As Boolean) As Boolean
Dim xlapp As Excel.Application
Dim wb As Workbook
Dim file_format As Long
'set file format according to the typ_a
If typ_a = "xlsx" Then
    file_format = 51 'xlOpenXMLWorkbook
ElseIf typ_a = "xls" Then
    file_format = 56 'xlExcel8
ElseIf typ_a = "csv" Then
    file_format = 6 'xlCSV
Else
    file_format = 51
End If

If output_file = "" Then
    output_file = Left(input_file, InStr(1, input_file, ".") - 1) & typ_a
End If
Set xlapp = xlappcreate(False, False)
Set wb = xlapp.Workbooks.Open(input_file)
wb.CheckCompatibility = False
wb.SaveAs filename:=output_file, FileFormat:=file_format
wb.Close savechanges:=True
If del = True Then
    kill input_file
End If
xlstyp = True
End Function

'Change all excel file types in the specific folders
Public Function xlstypall(ByVal fdr_path As String, ByVal typ_b As String, ByVal typ_a As String) As Boolean
Dim xlapp As Excel.Application
Dim filename As String
Dim wb As Workbook
Dim FullName As String
Dim targetname As String
Dim file_format As Long

Set xlapp = xlappcreate(False, False)
If Right(fdr_path, 1) <> "\" Then
    fdr_path = fdr_path & "\"
End If

'set file format according to the typ_a
If typ_a = "xlsx" Then
    file_format = 51 'xlOpenXMLWorkbook
ElseIf typ_a = "xls" Then
    file_format = 56 'xlExcel8
ElseIf typ_a = "csv" Then
    file_format = 6 'xlCSV
Else
    file_format = 51
End If

filename = dir(fdr_path)

Do While filename <> ""
    If Right(filename, Len(filename) - InStrRev(filename, ".")) = typ_b Then
        FullName = fdr_path & filename
        targetname = fdr_path & Left(filename, InStr(filename, ".")) & typ_a
        Set wb = xlapp.Workbooks.Open(FullName)
        wb.CheckCompatibility = False
        wb.SaveAs filename:=targetname, FileFormat:=file_format
        wb.Close savechanges:=True
        kill FullName
        xlstypall = True
    End If
    filename = dir()
Loop
End Function

'if worksheet exists
Public Function wsex(ByVal ws As Worksheet) As Boolean
Dim str As String
On Error Resume Next
str = ws.name
If ERR = 0 Then
    wsex = True
ElseIf ERR > 0 Then
    wsex = False
End If
On Error GoTo 0
End Function

'Whether any worksheet exists or not (search by name)
Public Function wsexn(ByVal ws_name As String, Optional ByRef wb As Workbook, Optional ByVal con As Boolean) As Boolean
Dim wba As Variant, wbb As Workbook
Dim ws As Worksheet
Dim wbs As Variant
Dim ex As Boolean
Dim b As Boolean
ex = wbex(wb)
If ex = False Then
    wbs = getwbs()
    For Each wba In wbs
        Set wbb = wba
        If wsexn(ws_name, wbb, con) = True Then
            wsexn = True
            Exit Function
        End If
    Next wba
Else
    For Each ws In wb.Worksheets
        If con = False Then
            b = (ws.name = ws_name)
        Else
            b = (InStr(1, ws.name, ws_name) <> 0)
        End If
        If b = True Then
            wsexn = True
            Exit Function
        End If
    Next ws
End If
End Function

'Set worksheet of a workbook
Public Function wsset(ByVal ws_name As String, ByVal wb As Workbook, Optional ByVal con As Boolean) As Worksheet
Dim ws As Worksheet
Dim b As Boolean
For Each ws In wb.Worksheets
    If con = False Then
        b = (ws.name = ws_name)
    Else
        b = (InStr(1, ws.name, ws_name) <> 0)
    End If
    If b = True Then
        Set wsset = ws
        Exit Function
    End If
Next ws
End Function

'Rename worksheet
Public Function wsrename(ByVal ws As Worksheet, ByVal name As String, Optional ByVal ovrwt As Boolean) As Worksheet
Dim wb As Workbook
If wsex(ws) = True Then
    wb = ws.Parent
    If wsex_b(name, wb) = False Then
        ws.name = name
    Else
        If ovrwt = True Then
            wb.Worksheets(name).Delete
            ws.name = name
        ElseIf ovrwt = False Then
            debug_err "wsrename", , "Worksheet named - " & name & " exists."
        End If
    End If
ElseIf wsex(ws) = False Then
    debug_err "wsrename", , "Worksheet not exists."
End If
End Function

'Count all active worksheets
Public Function wsct() As Long
Dim d As Integer
Dim i As Long
Dim wb As Variant, wbs As Variant
Dim ws As Worksheet
Dim ct As Long
wbs = getactwb()
For Each wb In wbs
    For Each ws In wb.Worksheets
        ct = ct + 1
    Next ws
Next wb
wsct = ct
End Function

'Count all active worksheets by excel applications, workbooks
Public Function wscts() As Variant
Dim d As Integer
Dim i As Long, j As Long, k As Long
Dim wb As Workbook, wbs() As Workbook
Dim ws As Worksheet
Dim ct As Long
Dim cts() As Variant
ct = wbct()
If ct > 0 Then
    wbs = getactwb
    d = aydim(wbs)
    If d = 1 Then
        ReDim cts(1 To UBound(wbs, 1))
        For j = 1 To UBound(wbs, 1)
            For Each ws In wbs(j).Worksheets
                cts(j) = cts(j) + 1
            Next ws
        Next j
    ElseIf d = 2 Then
        ReDim cts(1 To UBound(wbs, 1), 1 To UBound(wbs, 2))
        For i = 1 To UBound(wbs, 1)
            For j = 1 To UBound(wbs, 2)
                For Each ws In wbs(i, j).Worksheets
                    cts(i, j) = cts(i, j) + 1
                Next ws
            Next j
        Next i
    End If
    wscts = cts
ElseIf ct = 0 Then
    debug_err "wscts", , "There is no active worksheets."
End If
End Function

'Get all active worksheets
Public Function getactws() As Variant
Dim d As Long
Dim i As Long, j As Long, k As Long
Dim wb As Workbook, wbs() As Workbook
Dim ws As Worksheet, wss() As Worksheet
Dim ct As Long
Dim cts As Variant
Dim v As Variant
ct = aymax(aytr1(wscts()))
wbs = getactwb()
d = aydim(wbs)
If d = 1 Then
    ReDim wss(1 To UBound(wbs, 1), 1 To ct)
    For i = 1 To UBound(wbs, 1)
        Set wb = wbs(i)
        k = 1
        For Each ws In wb.Worksheets
            Set wss(i, k) = ws
            k = k + 1
        Next ws
    Next i
ElseIf d = 2 Then
    ReDim wss(1 To UBound(wbs, 1), 1 To UBound(wbs, 2), 1 To ct)
    For i = 1 To UBound(wbs, 1)
        For j = 1 To UBound(wbs, 2)
            Set wb = wbs(i, j)
            k = 1
            For Each ws In wb.Worksheets
                Set wss(i, j, k) = ws
                k = k + 1
            Next ws
        Next j
    Next i
End If
getactws = wss
End Function

'Count of worksheets of a workbook
Public Function wsct_wb(ByVal wb As Workbook) As Long
Dim i As Long
Dim ws As Worksheet
For Each ws In wb.Worksheets
    i = i + 1
Next ws
wsct_wb = i
End Function

'Get all worksheets of a workbook
Public Function getws_wb(ByVal wb As Workbook) As Variant
Dim ws As Worksheet, wss() As Worksheet
Dim ct As Long
ct = wsct_wb(wb)
ReDim wss(1 To ct) As Worksheet
i = 1
For Each ws In wb.Worksheets
    Set wss(i) = ws
    i = i + 1
Next ws
getws_wb = wss
End Function

'Whether a worksheet is empty
Public Function ws_empty(ByVal ws As Worksheet) As Boolean
If ws.Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
    ws_empty = True
Else
    ws_empty = False
End If
End Function

'Number of non empty cells in a worksheet
Public Function wscta(ByVal ws As Worksheet) As Long
    wscta = ws.Application.WorksheetFunction.CountA(ws.Cells)
End Function

'Dimension of the worksheets
Public Function wsrge(ByVal ws As Worksheet) As Variant
Dim rge As Range
Dim A(1 To 2) As Variant
Set rge = ws.UsedRange
A(1) = rge.Rows.Count
A(2) = rge.Columns.Count
wsrge = A
End Function

'Dimension of a workbook
Public Function wbrge(ByVal wb As Workbook) As Variant
Dim RS As Collection
Dim cs As Collection
Dim k As Long
Dim R As Long
Dim c As Long
Dim A As Variant, A1 As Variant
Dim ws As Worksheet, ws1 As Worksheet
Set RS = New Collection
Set cs = New Collection
ReDim A(0 To 2) As Variant
For Each ws In wb.Worksheets
    k = k + 1
    A1 = wsrge(ws)
    RS.add A1(1)
    cs.add A1(2)
Next ws
A(0) = k
A(1) = aymax(CltAy(RS))
A(2) = aymax(CltAy(cs))
wbrge = A
End Function

'Hide or unhide worksheet and return the current status (1 Hide, 2 Unhide)
Public Function Hiws(ByVal ws As Worksheet, Optional ByVal hide As Byte) As Boolean
    If hide = 0 Then
        If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
            Hiws = True
        ElseIf ws.Visible = xlSheetVisible Then
            Hiws = False
        End If
    ElseIf hide = 1 Then
        If ws.Visible <> xlSheetHidden Then
            On Error Resume Next
            ws.Visible = xlSheetHidden
            If ERR.Number > 0 Then
                debug_err "hiws", , "The worksheet cannot be hidden"
                Exit Function
            End If
        End If
        Hiws = True
    ElseIf hide = 2 Then
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
        Hiws = False
    Else
        debug_err "Hiws", , "Hide can only be either 0, 1 or 2"
    End If
End Function

'show or unshow scroll bar of the excel (1 Hide, 2 Unhide)
Public Function hiscroll(Optional ByVal Hori As Boolean, Optional hide As Byte) As Boolean
With ActiveWindow
    If hide = 0 Then
        If Hori = False Then
            hiscroll = .DisplayVerticalScrollBar
        ElseIf Hori = True Then
            hiscroll = .DisplayHorizontalScrollBar
        End If
    If hide = 1 Then
        If Hori = False Then
            If .DisplayVerticalScrollBar = True Then
                .DisplayVerticalScrollBar = False
            End If
        ElseIf Hori = True Then
            If .DisplayHorizontalScrollBar = True Then
                .DisplayHorizontalScrollBar = False
            End If
        End If
    ElseIf hide = 2 Then
        If Hori = False Then
            If .DisplayVerticalScrollBar = False Then
                .DisplayVerticalScrollBar = True
            End If
        ElseIf Hori = True Then
            If .DisplayHorizontalScrollBar = False Then
                .DisplayHorizontalScrollBar = True
            End If
        End If
    End If
End With
End Function

'Protect or unprotect workbook
'Protect: 1 protect, 2 unprotect
Public Function Protectwb(ByVal wb As Workbook, Optional ByVal protect As Byte, Optional ByVal Pw As String) As Boolean
    If protect = 0 Then
        Protectwb = wb.ProtectStructure
    ElseIf protect = 1 Then
        On Error GoTo err1
        wb.protect Structure:=True, Windows:=False, password:=Pw
        Exit Function
err1:
        debug_err "Protectwb", , "Workbook could not be protected"
        Exit Function
    ElseIf protect = 2 Then
        On Error GoTo err2
        wb.Unprotect password:=Pw
        Exit Function
err2:
        debug_err "Protectwb", , "Workbook could not be protected. (Password incorrect)"
        Exit Function
    End If
Protectwb = wb.ProtectStructure
End Function

'Workbook to array (3 dimension)
Public Function wbay(ByVal wb As Workbook) As Variant
Dim i As Long, j As Long, k As Long
Dim A() As Variant, AA As Variant
Dim ws As Variant
Dim R As Variant
R = wbran(wb)
ReDim A(1 To R(0), 0 To R(1), 0 To R(2)) As Variant
k = 1
For Each ws In wb.Worksheets
    A(k, 0, 0) = ws.name
    AA = wsay(ws)
    If IsArray(AA) = True Then
        For i = LBound(AA, 1) To UBound(AA, 1)
            For j = LBound(AA, 2) To UBound(AA, 2)
                A(k, i, j) = AA(i, j)
            Next j
        Next i
    End If
    k = k + 1
Next ws
wbay = A
End Function

'Worksheet to array (2 dimension)
Public Function wsay(ByVal ws As Worksheet) As Variant
wsay = ws.UsedRange
End Function

'Select columns of worksheet to array (2 dimension)
Public Function wsayc(ByVal ws As Worksheet, ByRef cols As Variant) As Variant
Dim i As Long, j As Long
Dim d As Long
Dim low As Long, up As Long
Dim A As Variant, A1 As Variant
Dim R As Long, c As Long
R = wsrge(ws)(1)
c = wsrge(ws)(2)
d = aydim(cols)
If d = 0 Then
    A = ws.Range(ws.Cells(1, cols), ws.Cells(R, cols))
    ReDim A1(1 To R) As Variant
    For i = 1 To R
        A1(i) = A(i, 1)
    Next i
    wsayc = A1
ElseIf d = 1 Then
    low = LBound(cols, 1)
    up = UBound(cols, 1)
    For i = low To up
        If cols(i) < 1 Or cols(i) > c Then
            debug_err "wsayc", , "At least one of the selected columns is out of the range of the worksheet, please check!"
            Exit Function
        End If
    Next i
    ReDim A(1 To R, 1 To (up - low + 1)) As Variant
    For j = low To up
        A1 = ws.Range(ws.Cells(1, cols(j)), ws.Cells(R, cols(j)))
        For i = 1 To R
            A(i, (j - low + 1)) = A1(i, 1)
        Next i
    Next j
    wsayc = A
Else
    debug_err "wsayc", , "The dimension of the cols is greater than one, please check!"
End If
End Function

'Worksheet sort
Function wssort(ByVal ws As Worksheet, ByVal cols As Variant, Optional ByVal hdr As Boolean, Optional ByVal sortop As Variant, Optional ord As Variant, Optional dataop As Variant) As Boolean
Dim b As Boolean, b1 As Boolean
Dim h As Integer
Dim R As Variant
Dim RS As Range
Dim op(1 To 3, 1 To 2) As Variant
b = wsex(ws)
If b = True Then
    b = isay(cols)
    R = wsrge(ws)
    Set RS = ws.Range(ws.Cells(1, 1), ws.Cells(R(1), R(2)))
    op(1, 1) = sortop
    op(2, 1) = ord
    op(3, 1) = dataop
    If IsMissing(op(1, 1)) = True Then op(1, 2) = xlSortOnValues
    If IsMissing(op(2, 1)) = True Then op(2, 2) = xlAscending
    If IsMissing(op(3, 1)) = True Then op(3, 2) = xlSortNormal
    If hdr = True Then
        h = xlYes
    Else
        h = xlNo
    End If
    If b = False Then
        b = IsNumeric(cols)
        If b = True Then
            ws.Sort.SortFields.add Key:=Columns(cols), sorton:=op(1, 2), Order:=op(2, 2), DataOption:=op(3, 2)
            With ws.Sort
                .SetRange RS
                .header = h
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        Else
            debug_err "wssort", "nnum"
        End If
    Else
        b = isay(cols, , True, 1)
        If b = True Then
            For i = LBound(cols, 1) To UBound(cols, 1)
                For j = 1 To 3
                    If isay(op(j, 1), , True, 1) = True Then
                        op(j, 2) = op(j, 1)(i)
                    ElseIf isay(ord) = False Then
                        op(j, 2) = op(j, 1)
                    End If
                Next j
                ws.Sort.SortFields.add Key:=Columns(cols(i)), sorton:=op(1, 2), Order:=op(2, 2), DataOption:=op(3, 2)
            Next i
            With ws.Sort
                .SetRange RS
                .header = h
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        Else
            debug_err "wssort", "nnum"
        End If
    End If
Else
    debug_err "wssort", "wsnex"
End If
End Function

'Worksheet Comparison, highlight the unequal cells
Public Function wscomp(ByVal ws As Worksheet, ByVal wsc As Worksheet, Optional ByVal color As Long) As Boolean
Dim i As Long, j As Long
Dim A As Variant, AC As Variant
A = wsay(ws)
AC = wsay(wsc)
If color = 0 Then color = 65535
For i = 1 To UBound(A, 1)
    For j = 1 To UBound(A, 2)
        If A(i, j) <> AC(i, j) Then
            ws.Cells(i, j).Interior.color = color
        End If
    Next j
Next i
End Function

'is a range
Public Function isrge(ByVal ran As Variant) As Boolean
Dim R As Range
'On Error Resume Next
Set R = ran
If ERR.Number = 0 Then
    isrge = True
Else
    isrge = False
End If
End Function

'range of the range
Public Function rgerge(ByVal ran As Range) As Variant
Dim R(1 To 2, 1 To 2) As Long
R(1, 1) = ran.row
R(2, 1) = R(1, 1) + ran.Rows.Count - 1
R(1, 2) = ran.column
R(2, 2) = R(1, 2) + ran.Columns.Count - 1
rgerge = R
End Function

'Range to array
Public Function rgeay(ByVal ran As Range, Optional ByVal d As Boolean) As Variant
Dim A As Variant, A1() As Variant
Dim i As Long
Dim R As Long, c As Long
R = ran.Rows.Count
c = ran.Columns.Count
ReDim A(1 To R, 1 To c) As Variant
A = ran
If d = False Then
    If R > 1 And c > 1 Then
        rgeay = A
    ElseIf R = 1 And c > 1 Then
        ReDim A1(1 To c)
        For i = 1 To c
            A1(i) = A(1, i)
        Next i
        rgeay = A1
    ElseIf R > 1 And c = 1 Then
        ReDim A1(1 To R)
        For i = 1 To R
            A1(i) = A(i, 1)
        Next i
        rgeay = A1
    ElseIf R = 1 And c = 1 Then
        rgeay = A
    End If
Else
    rgeay = ran
End If
End Function

'Ranges to array (2-dimensional) (S: Set Statment in SAS, M: Merge Statement in SAS)
Public Function rgesay(ByVal ran As Variant, ByVal SM As String) As Variant
Dim b As Boolean
Dim A As Variant
Dim i As Long, j As Long, k As Long
Dim R As Variant, RS As Variant, rt(1 To 2) As Variant
Dim RV As Variant, F As Variant
Dim AR As Long, AC As Long
Dim S As Long
b = isay(ran, , , 1)
If b = True Then
    R = ayrge(ran)
    ReDim RS(R(1, 1) To R(2, 1)) As Variant
    ReDim RV(R(1, 1) To R(2, 1)) As Variant
    For k = R(1, 1) To R(2, 1)
        rt(1) = ran(k).Rows.Count
        rt(2) = ran(k).Columns.Count
        RS(k) = rt
        RV(k) = rgeay(ran(k), True)
    Next k
    If SM = "M" Then
        AR = aymax(jagtr1(RS, Array(1)))
        AC = aysum(jagtr1(RS, Array(2)))
        ReDim F(1 To AR, 1 To AC) As Variant
        S = 0
        For k = R(1, 1) To R(2, 1)
            For i = 1 To RS(k)(1)
                For j = 1 To RS(k)(2)
                    F(i, S + j) = RV(k)(i, j)
                Next j
            Next i
            S = S + RS(k)(2)
        Next k
        rgesay = F
    ElseIf SM = "S" Then
        AR = aysum(jagtr1(RS, Array(1)))
        AC = aymax(jagtr1(RS, Array(2)))
        ReDim F(1 To AR, 1 To AC) As Variant
        S = 0
        For k = R(1, 1) To R(2, 1)
            For i = 1 To RS(k)(1)
                For j = 1 To RS(k)(2)
                    F(S + i, j) = RV(k)(i, j)
                Next j
            Next i
            S = S + RS(k)(1)
        Next k
        rgesay = F
    End If
Else
    debug_err "rgesay", , "ran must be a 1-dimensional array."
End If
End Function

'Range's CountA
Public Function rgecta(ByVal ran As Range) As Long
    rgecta = ran.Application.WorksheetFunction.CountA(ran)
End Function

'Ranges Equal or not
Public Function rgeseq(ByVal ran1 As Range, ByVal ran2 As Range) As Boolean
Dim A As Variant, b As Variant
A = rgeay(ran1)
b = rgeay(ran2)
If ayequal(A, b) = True Then
    rgeseq = True
End If
End Function

'Copy range values
Function copyrangevalues(ByVal ran_f As Range, ByVal ran_t As Range, Optional ByVal overwrite As Boolean) As Boolean
Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
Dim v As Long
r1 = ran_f.Rows.Count
c1 = ran_f.Columns.Count
r2 = ran_t.Rows.Count
c2 = ran_t.Columns.Count
If r1 <> r2 Or c1 <> c2 Then
    debug_err "copyrange", , "The dimensions of the ranges do not match."
    Exit Function
End If
If overwrite = False Then
    If rgecta(ran_t) > 0 Then
        v = MsgBox("Warning! Are you sure to overwrite the values in the range?", vbYesNo, "Warning!")
        If v = vbNo Then
            Exit Function
        End If
    End If
End If
ran_t.Value = ran_f.Value
copyrangevalues = True
End Function

'Open workbook
Public Function wbopen(ByVal path As String, Optional xlapp As Application) As Workbook
Dim b As Boolean
Dim wb As Workbook
b = appex(xlapp)
If b = False Then
    Set xlapp = GetObject(, "Excel.Application")
End If
On Error GoTo err_handle
Set wb = xlapp.Workbooks.Open(path)
wbopen = True
err_handle:
If ERR.Number = 1004 Then
    MsgBox "The path is invalid. Please check"
End If
End Function

'Add n Workbooks
Public Function wbadd(Optional ByVal n As Long, Optional ByVal save As Boolean, Optional ByVal path As String, Optional ByVal name As Variant, Optional ByVal format As String, Optional ByVal hide As Boolean, Optional ByVal clse As Boolean) As Workbook
Dim i As Long
Dim d As Long
Dim path1 As String
Dim xlapp As Excel.Application
Dim wb As Workbook, twb As Workbook
Dim low As Long, up As Long
Set twb = ThisWorkbook
If n = 0 Then n = 1
If hide = False Then
    On Error Resume Next
    Set xlapp = GetObject(, "Excel.Application")
    If xlapp Is Nothing Then
        Set xlapp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0
Else
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Visible = Not (hide)
End If
xlapp.Application.DisplayAlerts = False
If path <> "" Then
    If Right(path, 1) <> "\" Then
        path = path & "\"
    ElseIf Right(path, 1) = "\" Then
        path = path
    End If
Else
    path = twb.path & "\"
End If
For i = 1 To n
    Set wb = xlapp.Workbooks.add
    If save = True Then
        path1 = path
        If format = "" Then format = "xlsx"
        If IsMissing(name) = False Then
            d = aydim(name)
            If d = 0 Then
                If name = "" Then name = "Workbook"
                If i = 1 Then
                    path1 = path1 & name & "." & format
                ElseIf i > 1 Then
                    path1 = path1 & name & "(" & i & ")." & format
                End If
            ElseIf d = 1 Then
                low = LBound(name, 1)
                up = UBound(name, 1)
                If n <= (up - low + 1) Then
                    path1 = path1 & name(low - 1 + i) & "." & format
                ElseIf n > (up - low + 1) Then
                    debug_err "wbadd", , "The number of items in the array of name is smaller than the number of worksheets to be added"
                    xlapp.Quit
                    Exit Function
                End If
            ElseIf d > 1 Then
                debug_err "wbadd", , "name should be a value or a 1-dimension array only. Please check"
                xlapp.Quit
                Exit Function
            End If
            wb.SaveAs path1
        Else
            name = "Workbook"
            If i = 1 Then
                path1 = path1 & name & "." & format
            ElseIf i > 1 Then
                path1 = path1 & name & "(" & i & ")." & format
            End If
            wb.SaveAs path1
        End If
    End If
    If clse = True Then
        wb.Close False
    End If
Next i
If clse = False Then
    Set wbadd = wb
Else
    xlapp.Quit
End If
End Function

'Add n worksheets (return the last worksheet added)
Public Function wsadd(ByVal wb As Workbook, Optional ByVal n As Long, Optional ByVal name As Variant, Optional ByVal ovrwt As Boolean, Optional ByVal Before As Worksheet, Optional ByVal After As Worksheet, Optional ByVal wsr As Long, Optional ByVal wsc As Long, Optional ByVal Ht As Variant, Optional ByVal Wt As Variant) As Worksheet
Dim i As Long, j As Long
Dim d As Long
Dim ws As Worksheet, ws1 As Worksheet
Dim R As Variant
Dim wsrp As Long, wscp As Long
If n = 0 Then n = 1
Application.DisplayAlerts = False
ERR.Clear
On Error Resume Next
For i = 1 To n
    If wsex(Before) = False And wsex(After) = False Then
        Set ws = wb.Worksheets.add(After:=wb.Sheets(wb.Sheets.Count))
    ElseIf wsex(Before) = True And wsex(After) = False Then
        Set ws = wb.Worksheets.add(Before:=Before)
    ElseIf wsex(Before) = False And wsex(After) = True Then
        Set ws = wb.Worksheets.add(After:=ws1)
        Set After = ws
    ElseIf wsex(Before) = True And wsex(After) = True Then
        debug_err "wsadd", , "Either Before or After must be null"
    End If
    If IsMissing(name) = False Then
        d = aydim(name)
        If d = 0 Then
            If i = 1 Then
                ws.name = name
                If ERR.Number = 1004 Then
                    If ovrwt = False Then
                        ws.Delete
                        debug_err "wsadd", , "The worksheet - " & name & " exists. Please check."
                        Exit Function
                    Else
                        wb.Sheets(name).Delete
                        ws.name = name
                    End If
                    ERR.Clear
                End If
            ElseIf i > 1 Then
                ws.name = name & "(" & i & ")"
                If ERR.Number = 1004 Then
                    If ovrwt = False Then
                        ws.Delete
                        debug_err "wsadd", , "The worksheet - " & name & "(" & i & ")" & " exists. Please check."
                        Exit Function
                    Else
                        wb.Sheets(name & "(" & i & ")").Delete
                        ws.name = name & "(" & i & ")"
                    End If
                    ERR.Clear
                End If
            End If
        ElseIf d = 1 Then
            R = ayrge(name)
            If n <= (R(2, 1) - R(1, 1) + 1) Then
                ws.name = name(R(1, 1) - 1 + i)
                If ERR.Number = 1004 Then
                    If ovrwt = False Then
                        ws.Delete
                        debug_err "wsadd", , "The worksheet - " & name(R(1, 1) - 1 + i) & " exists. Please check."
                        Exit Function
                    ElseIf ovrwt = True Then
                        wb.Sheets(name(R(1, 1) - 1 + i)).Delete
                        ws.name = name(R(1, 1) - 1 + i)
                    End If
                    ERR.Clear
                End If
            ElseIf n > (R(2, 1) - R(1, 1) + 1) Then
                debug_err "wsadd", , "The number of items in the array of name is smaller than the number of worksheets to be added"
                Exit Function
            End If
        Else
            debug_err "wsadd", , "name should be a value or a 1-dimension array only. Please check"
            Exit Function
        End If
    Else
    End If
    If wsr > 0 Then
        wsrp = ws.Rows.Count
        ws.Range(Rows(wsr + 1), Rows(wsrp)).Hidden = True
    ElseIf wsr < 0 Then
        debug_err "wsadd", , "wsr must be greater than 0."
    End If
    If wsc > 0 Then
        wscp = ws.Columns.Count
        ws.Range(Columns(wsc + 1), Columns(wscp)).Hidden = True
    ElseIf wsc < 0 Then
        debug_err "wsadd", , "wsc must be greater than 0."
    End If
Next i
Set wsadd = ws
End Function

'Delete Worksheets
Public Function wsdel(ByVal wb As Workbook, ByVal wsname As String, Optional ByVal con As Boolean) As Boolean
Dim ws As Worksheet
For Each ws In wb.Worksheets
    If con = False Then
        If ws.name = wsname Then
            ws.Delete
            wsdel = True
        End If
    Else
        If InStr(1, ws.name, wsname) <> 0 Then
            ws.Delete
            wsdel = True
        End If
    End If
Next ws
End Function

'Write an array to Worksheet (1 or 2 dimension)
Public Function ayws(ByVal ay As Variant, ByVal ws As Worksheet, Optional ByVal RC As String, Optional ByVal row As Long, Optional ByVal column As Long, Optional ByVal overwrite As Boolean) As Boolean
Dim b As Boolean, b1 As Boolean
Dim i As Long
Dim d As Integer
Dim R As Variant
Dim A As Variant
If overwrite = False Then
    b = ws_empty(ws)
    If b = False Then
        b1 = MsgBox("The worksheet you are going to write is not empty, are you sure to overwrite it?", vbYesNo)
        If b1 = vbNo Then
            Exit Function
        End If
    End If
End If
d = aydim(ay)
R = ayrge(ay)
If RC = "" Then
    If row = 0 Then
        row = 1
    ElseIf row < 0 Or row > (1048576 - R(2, 1)) Then
        debug_err "ayws", , "The number of row exceeds range"
        Exit Function
    End If
    If column = 0 Then
        column = 1
    ElseIf column < 0 Or column > (16384 - R(2, 2)) Then
        debug_err "ayws", , "The number of column exceeds range"
        Exit Function
    End If
End If
If d = 0 Then
    ws.Cells(row, column).Value = ay
ElseIf d = 1 Then
    If RC = "" Then
        RC = "R"
    End If
    If RC = "R" Then
        ReDim A(R(1, 1) To R(2, 1), 1 To 1) As Variant
        For i = R(1, 1) To R(2, 1)
            A(i, 1) = ay(i)
        Next i
        ws.Range(ws.Cells(row, column), ws.Cells(row + (R(2, 1) - R(1, 1)), column)) = A
    ElseIf RC = "C" Then
        ReDim A(1 To 1, R(1, 1) To R(2, 1)) As Variant
        For i = R(1, 1) To R(2, 1)
            A(1, i) = ay(i)
        Next i
        ws.Range(ws.Cells(row, column), ws.Cells(row, column + (R(2, 1) - R(1, 1)))) = A
    Else
        debug_err "ayws", , "Please specify row(R) or column(C)"
        Exit Function
    End If
ElseIf d = 2 Then
    ws.Range(ws.Cells(row, column), ws.Cells(row + (R(2, 1) - R(1, 1)), column + (R(2, 2) - R(1, 2)))) = ay
ElseIf d > 2 Then
    debug_err "ayws", , "The dimension of the array exceeds 2, please check."
End If
End Function
    
'Find position of the worksheet (Maximum 2 conditions)
'parameters: (RC = False: search column, RC = True: search row) (P: row / column number to be found)
Public Function wspos(ByVal ws As Worksheet, ByVal str1 As String, Optional ByVal RC As Boolean, Optional ByVal P As Long, Optional ByVal oper1 As String, Optional ByVal fun1 As String, Optional ByVal par1 As Integer, Optional ByVal str2 As String, Optional ByVal oper2 As String, Optional ByVal fun2 As String, Optional ByVal par2 As Integer) As Long
Dim S As String
Dim i As Long
On Error Resume Next
If P = 0 Then
    P = 1
End If
If oper1 = "" Then
    oper1 = "="
End If
i = 1
S = wspos_c(ws, i, RC, P, oper1, str1, fun1, par1)
Do Until Evaluate(S)
    If ERR.Number > 0 Then
        debug_err "wspos", , "Position not found for " & str1 & "."
        Exit Function
    End If
    i = i + 1
    S = wspos_c(ws, i, RC, P, oper1, str1, fun1, par1)
Loop
If str2 = "" Then
    wspos = i
ElseIf str2 <> "" Then
    If oper2 = "" Then
        oper2 = "="
    End If
    S = wspos_c(ws, i, RC, P, oper2, str2, fun2, par2)
    Do Until Evaluate(S)
        If ERR.Number > 0 Then
            debug_err "wspos", , "Position not found."
            Exit Function
        End If
        i = i + 1
        S = wspos_c(ws, i, RC, P, oper2, str2, fun2, par2)
    Loop
    wspos = i - 1
End If
End Function
    Public Function wspos_c(ByVal ws As Worksheet, ByVal i As Long, ByVal RC As Boolean, ByVal P As Long, ByVal oper As String, ByVal str As String, ByVal Fun As String, ByVal par As Integer) As String
    Dim st As String
    Dim S As Variant
    If RC = False Then
        S = ws.Cells(P, i).Value
    ElseIf RC = True Then
        S = ws.Cells(i, P).Value
    End If
    st = CStr(chr(34))
        If Fun = "" And par = 0 Then
            wspos_c = st & CStr(S) & st & oper & st & str & st
        ElseIf Fun = "Left" Then
            If par > 0 Then
                wspos_c = st & CStr(Left(S, par)) & st & oper & st & str & st
            ElseIf par <= 0 Then
                debug_err "ficol_cel"
            End If
        ElseIf Fun = "Right" Then
            If par > 0 Then
                wspos_c = st & CStr(Right(S, par)) & st & oper & st & str & st
            ElseIf par <= 0 Then
                debug_err "ficol_cel"
            End If
        ElseIf Fun = "Instr" Then
            wspos_c = CStr(InStr(1, S, str)) & "<> 0"
        Else
            debug_err "wspos_c"
        End If
    End Function
'End of Find Column of the worksheet by row

'Get Font List of Excel
Function FontList() As Variant
Dim i As Long, c As Long
Dim FontNamesCtrl As CommandBarControl
Dim l As Variant
Set FontNamesCtrl = Application.CommandBars("Formatting").FindControl(id:=1728)
c = FontNamesCtrl.ListCount
ReDim l(1 To c) As Variant
For i = 1 To c
    l(i) = FontNamesCtrl.List(i)
Next i
FontList = aysort(ayuni(l))
End Function

'Border line of range in the excel
Public Function rge_bdl(ByVal rge As Range, Optional ByVal linestyle As Long, Optional ByVal weight As Long, Optional out As Boolean, Optional vert As Boolean, Optional horizon As Boolean) As Boolean
If linestyle = 0 Then linestyle = xlContinuous
If weight = 0 Then weight = xlThin
If out = False And vert = False And horizon = False Then
    For i = 1 To 6
        With rge.Borders(BDL(i))
            .linestyle = linestyle
            .weight = weight
        End With
    Next i
ElseIf out = True Then
    For i = 1 To 4
        With rge.Borders(BDL(i))
            .linestyle = linestyle
            .weight = weight
        End With
    Next i
End If
If vert = True Then
    With rge.Borders(BDL(5))
        .linestyle = linestyle
        .weight = weight
    End With
End If
If horizon = True Then
    With rge.Borders(BDL(6))
        .linestyle = linestyle
        .weight = weight
    End With
End If
rge_bdl = True
End Function
    Public Function BDL(ByVal n As Integer) As Long
        BDL = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)(n - 1)
    End Function

'Default useful formats for range of the excel
'1: String format
'2: Bold
'3: Italic
'4: Font name
Public Function rge_fmt(ByVal rge As Range, Optional ByVal c As Long, Optional ByVal font_name As String) As Boolean
If c = 0 Then c = 1
If c = 1 Then
    rge.NumberFormatLocal = "@"
ElseIf c = 2 Then
    rge.Font.Bold = -1
ElseIf c = 3 Then
    rge.Font.Italic = -1
ElseIf c = 4 Then
    If font_name <> "" Then
        rge.Font.name = font_name
    ElseIf font_name = "" Then
        debug_err "rgefmt", , "Please specify the font name."
        Exit Function
    End If
Else
    debug_err "rgefmt", , "Please sepecify the correct default format."
End If
End Function

'For absolute value converting number format
Function convert_numberformat(ByVal rge As Range, ByVal proformat As String, Optional ByVal preformat As String, Optional min As Long, Optional max As Long) As Boolean
Dim v As Variant
If preformat = "" Then
    If min = 0 And max = 0 Then
        rge.NumberFormat = proformat
    Else
        For Each v In rge.Cells
            If Abs(v.Value) > min And Abs(v.Value) < max Then
                v.NumberFormat = proformat
            End If
        Next v
    End If
Else
    If min = 0 And max = 0 Then
        For Each v In rge.Cells
            If v.NumberFormat = preformat Then
                v.NumberFormat = proformat
            End If
        Next v
    Else
        For Each v In rge.Cells
            If v.NumberFormat = preformat Then
                If Abs(v.Value) > min And Abs(v.Value) < max Then
                    v.NumberFormat = proformat
                End If
            End If
        Next v
    End If
End If
convert_numberformat = True
End Function


'F!: Functions for access

'******************************************************************************************************************************************************

'Get all current database
Public Function Getdbs() As Variant
Dim i As Long
Dim dbs() As DAO.Database
Dim accapps As Variant
Dim ct As Long, ct1 As Long
ct = accappct()
accapps = Getaccapps()
ReDim dbs(1 To ct) As DAO.Database
For i = 1 To ct
    Set dbs(i) = accapps(i).CurrentDb
Next i
Getdbs = dbs
End Function

'table count of a database
Public Function tdfct(ByVal db As Database, Optional ByVal sys As Boolean) As Long
Dim i As Long: i = 1
Dim tdf As TableDef
For Each tdf In db.TableDefs
    If sys = False Then
        If Not (tdf.name Like "MSys*" Or tdf.name Like "~*") Then
            i = i + 1
        End If
    ElseIf sys = True Then
        i = i + 1
    End If
Next tdf
tdfct = i - 1
End Function
'
''table counts by database
Public Function tdfcts(Optional ByVal sys As Boolean) As Variant
Dim i As Long, j As Long
Dim dbs As Variant
Dim tdf As TableDef
Dim ct As Long
Dim cts() As Long
ct = accappct()
dbs = Getdbs()
ReDim cts(1 To ct) As Long
For i = 1 To ct
    For Each tdf In dbs(i).TableDefs
        If sys = False Then
            If Not (tdf.name Like "MSys*" Or tdf.name Like "~*") Then
                cts(i) = cts(i) + 1
            End If
        ElseIf sys = True Then
            cts(i) = cts(i) + 1
        End If
    Next tdf
Next i
tdfcts = cts
End Function

'Get all tables by database (incomplete)
Public Function gettdfs(Optional ByVal sys As Boolean) As Variant
Dim i As Long, j As Long
Dim ct As Long, cts As Variant
Dim dbs As Variant
Dim tdf As TableDef, tdfs() As Variant
ct = accappct()
dbs = Getdbs()
cts = tdfcts(sys)
ReDim tdfs(1 To ct, 1 To aymax(cts)) As TableDef
For i = 1 To ct
    j = 1
    For Each tdf In dbs(i).TableDefs
        If sys = False Then
            If Not (tdf.name Like "MSys*" Or tdf.name Like "~*") Then
                Set tdfs(i, j) = tdf
                j = j + 1
            End If
        ElseIf sys = True Then
            Set tdfs(i, j) = tdf
            j = j + 1
        End If
    Next tdf
Next i
gettdfs = tdfs
End Function

'Get all tables of a specify database
Public Function gettdfs_db(ByVal db As Database, Optional ByVal sys As Boolean) As Variant
Dim i As Long
Dim tdf As TableDef, tdfs() As TableDef
ct = tdfct(db, sys)
ReDim tdfs(1 To ct) As TableDef
i = 1
For Each tdf In db.TableDefs
    If sys = False Then
        If Not (tdf.name Like "MSys*" Or tdf.name Like "~*") Then
            Set tdfs(i) = tdf
            i = i + 1
        End If
    ElseIf sys = True Then
        Set tdfs(i) = tdf
        i = i + 1
    End If
Next tdf
gettdfs_db = tdfs
End Function

'Access table to array
Public Function tdfay(ByVal tdf As TableDef) As Variant
Dim i As Long, j As Long
Dim R As Long, c As Long
Dim A() As Variant
Dim RS As DAO.Recordset
Set RS = tdf.OpenRecordset
R = RS.RecordCount
c = RS.Fields.Count
ReDim A(1 To R, 1 To c) As Variant
For i = 1 To R
    For j = 1 To c
        A(i, j) = RS.Fields(j - 1).Value
    Next j
    RS.MoveNext
Next i
tdfay = A
End Function

'Is Access table empty
Public Function tdf_empty(ByVal tdf As TableDef) As Boolean
Dim RS As Recordset
Dim R As Long, c As Long
R = tdf.RecordCount
c = tdf.Fields.Count
If R = 0 Or c = 0 Then
    tdf_empty = True
Else
    tdf_empty = False
End If
End Function

'Array to Access table
Public Function aytdf(ByVal ay As Variant, ByVal db As Database, ByVal name As String, Optional ByVal hide As Boolean) As Variant
Dim i As Long, j As Long
Dim d As Long
Dim R As Variant
Dim tbl As DAO.TableDef
Dim fld As DAO.Field
Dim RS As DAO.Recordset
d = aydim(ay)
If d > 2 Then
    debug_err "aytdf", "2du"
    Exit Function
End If
If d >= 1 Then
    R = ayrge(ay)
End If
Set tbl = db.CreateTableDef(name, 1)
For j = R(1, 2) To R(2, 2)
    Set fld = tbl.CreateField("C" & j - R(1, 2) + 1, dbText, 20)
    tbl.Fields.Append fld
    If d < 2 Then Exit For
Next j
db.TableDefs.Append tbl
Set RS = db.OpenRecordset(name, dbOpenDynaset)
If d = 0 Then
    RS.AddNew
    RS!Field(0) = ay
    RS.Update
ElseIf d = 1 Then
    For i = R(1, 1) To R(2, 1)
        RS.AddNew
        RS.Fields("C1").Value = ay(i)
        RS.Update
    Next i
ElseIf d = 2 Then
    For i = R(1, 1) To R(2, 1)
        RS.AddNew
        For j = R(1, 2) To R(2, 2)
            RS.Fields("C" & j).Value = ay(i, j)
        Next j
        RS.Update
    Next i
End If
If hide = False Then
    tbl.Attributes = (tbl.Attributes - dbHiddenObject)
    RefreshDatabaseWindow
End If
End Function

'******************************************************************************************************************************************************
'G!: Functions for Word
'H!: Functions for Powerpoint
'I!: Functions for Outlook

'******************************************************************************************************************************************************
'If outlook Application exist
Function Outlookex() As Boolean
Dim outapp As Object
On Error Resume Next
Set outapp = GetObject(, "Outlook.Application")
If ERR.Number = 0 Then
    Outlookex = True
ElseIf ERR.Number = 429 Then
    Outlookex = False
End If
End Function

'Create Outlook Mail
Function CreateOutlookMail(ByVal subject As String, ByVal Body As String, ByVal towhom As String, Optional ByVal ccwhom As String, Optional ByVal attachment_path As String, Optional ByVal hide As Boolean) As Boolean
Dim oex As Boolean
Dim outapp As Object
Dim OutMail As Object
On Error Resume Next
Set outapp = GetObject(, "Outlook.Application")
If ERR.Number = 0 Then
    oex = True
ElseIf ERR.Number = 429 Then
    oex = False
End If
On Error GoTo 0
If oex = False Then
    Set outapp = CreateObject("Outlook.Application")
End If
Set OutMail = outapp.CreateItem(0)
With OutMail
    .subject = subject
    .To = towhom
    .Body = Body
    If ccwhom <> "" Then
        .CC = ccwhom
    End If
    If attachment_path <> "" Then
        .Attachments.add attachment_path
    End If
    If hide = False Then
        .Display
    End If
End With
Set OutMail = Nothing
Set outapp = Nothing
End Function

'******************************************************************************************************************************************************
'J!: Other functions

'Collection to Array
Public Function CltAy(Clt As Collection) As Variant
Dim i As Integer
Dim A() As Variant
Dim ct As Long
ct = Clt.Count
ReDim A(1 To ct) As Variant
For i = 1 To Clt.Count
    A(i) = Clt.Item(i)
Next
CltAy = A
End Function

'Counting run time of the Program(s)
Public Function timect(ByVal prog As Variant) As Double
Dim i As Long
Dim TS As Double, te As Double
Dim d As Long
Dim R As Variant
Dim ty As String, proerr As String
ty = aytyp(prog)
If ty = "String" Then
    d = aydim(prog)
    TS = Timer()
    On Error Resume Next
    If d = 0 Then
        Application.run prog
        If ERR.Number = 1004 Then GoTo SKip
    ElseIf d = 1 Then
        R = ayrge(prog)
        For i = R(1, 1) To R(2, 1)
            Application.run prog(i)
            If ERR.Number = 1004 Then GoTo SKip
        Next i
    ElseIf d > 1 Then
        debug_err "timect", "1du"
        GoTo skip1
    End If
    te = Timer()
    timect = te - TS
    Exit Function
ElseIf ty <> "String" Then
    debug_err "timect", , "The program name(s) must be string."
    GoTo skip1
End If
SKip:
    If d = 0 Then
        proerr = prog
    ElseIf d = 1 Then
        proerr = prog(i)
    End If
    debug_err "timect", , "The program - " & proerr & " do not exist."
skip1:
    timect = -1
End Function

'Msgbox of the run time of the Program(s)
Public Function timect_msg(ByVal prog As Variant, Optional conf As Long, Optional ByVal tst As Boolean) As Boolean
Dim t As Double
Dim fin_msg As String
t = timect(prog)
If t > 0 Then
    fin_msg = ""
    If conf = 0 Or conf = vbYes Then
        If fin_msg = "" Then
            fin_msg = "The job is done!"
        End If
        fin_msg = fin_msg & vbNewLine & vbNewLine & "It takes "
        If tst = False Then
            If Int(t / 60) = 0 Then
                fin_msg = fin_msg & Round(t, 1) & " seconds."
            ElseIf Int(t / 60) > 0 Then
                fin_msg = fin_msg & Int(t / 60) & " minutes and " & Int(t Mod 60) & " seconds."
            End If
        ElseIf tst = True Then
            fin_msg = fin_msg & t & " seconds."
        End If
        MsgBox fin_msg
    Else
        Exit Function
    End If
End If
End Function

'Ping the address
Function Ping(ByVal address As String) As String
Dim objPing As Object
Dim objStatus As Object
Dim Result As String
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("Select * from Win32_PingStatus Where Address = '" & address & "'")
For Each objStatus In objPing
    Select Case objStatus.StatusCode
        Case 0: Result = "Connected"
        Case 11001: Result = "Buffer too small"
        Case 11002: Result = "Destination net unreachable"
        Case 11003: Result = "Destination host unreachable"
        Case 11004: Result = "Destination protocol unreachable"
        Case 11005: Result = "Destination port unreachable"
        Case 11006: Result = "No resources"
        Case 11007: Result = "Bad option"
        Case 11008: Result = "Hardware error"
        Case 11009: Result = "Packet too big"
        Case 11010: Result = "Request timed out"
        Case 11011: Result = "Bad request"
        Case 11012: Result = "Bad route"
        Case 11013: Result = "Time-To-Live (TTL) expired transit"
        Case 11014: Result = "Time-To-Live (TTL) expired reassembly"
        Case 11015: Result = "Parameter problem"
        Case 11016: Result = "Source quench"
        Case 11017: Result = "Option too big"
        Case 11018: Result = "Bad destination"
        Case 11032: Result = "Negotiating IPSEC"
        Case 11050: Result = "General failure"
        Case Else: Result = "Unknown host"
    End Select
Next
Ping = Result
End Function

'Get Local IP
Public Function getip(Optional ByVal comname As String, Optional ByVal IPv6 As Boolean) As String
Dim objWMIservice As Object, IPConfigSet As Object, ipconfig As Variant
Dim IPAddress As Variant
If comname = vbNullString Then comname = "." '. stands for local computer name
'Connect to the WMI service
On Error Resume Next
Set objWMIservice = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & comname & "\root\cimv2")
If ERR.Number = 70 Then
    debug_err "getip", , "No authority for the computer - " & comname
    Exit Function
End If
'Get all TCP/IP-enabled network adapters
Set IPConfigSet = objWMIservice.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
'Get all IP addresses associated with these adapters
For Each ipconfig In IPConfigSet
    If Not IsNull(ipconfig.IPAddress) Then
        IPAddress = ipconfig.IPAddress
    ElseIf IsNull(ipconfig.IPAddress) Then
        getip = "This computer does not have any IP."
        Exit Function
    End If
Next
If IPv6 = False Then
    getip = CStr(IPAddress(0))
ElseIf IPv6 = True Then
    getip = CStr(IPAddress(1))
End If
End Function

'******************************************************************************************************************************************************

'J1!: Mouse & Keybroad Control Functions
'Left click
Public Function LeftClick(Optional n As Long, Optional ms As Long) As Long
If n = 0 Then
    n = 1
End If
For i = 1 To n
    mouse_event LEFTDOWN, 0, 0, 0, 0
    mouse_event LEFTUP, 0, 0, 0, 0
    If ms > 0 Then
        Pause ms
    End
Next i
LeftClick = n
End Function

'Right click
Public Function RightClick(Optional n As Long, Optional ms As Long) As Long
If n = 0 Then
    n = 1
End If
For i = 1 To n
    mouse_event RIGHTDOWN, 0, 0, 0, 0
    mouse_event RIGHTUP, 0, 0, 0, 0
    If ms > 0 Then
        Pause ms
    End
Next i
RightClick = n
End Function

'Move the mouse
Public Function MouseMove(x As Integer, y As Integer, Optional relative_move As Boolean) As Boolean
Dim x1 As Long
Dim y1 As Long
Dim CCoord As POINT
If relative_move = False Then
    If x >= 0 And y >= 0 Then
        SetCursorPOs x, y
    ElseIf x < 0 Or y < 0 Then
        debug_err "MouseMove", , "Both x and y must be greater than 0"
    End If
ElseIf relative_move = True Then
    GetCursorPos CCoord
    x1 = CCoord.Xcoord
    y1 = CCoord.Ycoord
    SetCursorPOs (x1 + x), (y1 + y)
End If
End Function

'Type the text
Public Function Sendtxt(txt As String) As String
SendKeys txt, True
Sendtxt = txt
End Function

'Get Color
Public Function GetColor(ByVal cursor As Boolean, Optional ByVal x As Integer, Optional ByVal y As Integer) As Long
Dim x1 As Long, y1 As Long
Dim color As Long
Dim lDC As Long
Dim CCoord As POINT
lDC = GetWindowDC(0)
If cursor = False Then
    GetColor = GetPixel(lDC, x, y)
Else
    If x = 0 And y = 0 Then
        GetCursorPos CCoord
        x1 = CCoord.Xcoord
        y1 = CCoord.Ycoord
        color = GetPixel(lDC, x1, y1)
    Else
        debug_err "GetColor", , "If Cursor = True, x and y should be 0 or blank, please check."
    End If
End If
If color <> -1 Then
    GetColor = color
Else
    debug_err "GetColor", , "Color got errors."
End If
End Function

'Return RGB of a color
Public Function ColorRGB(ByVal color As Long) As Variant
Dim c(1 To 3) As Long
If color <= 16777215 And color >= 0 Then
    c(1) = color Mod 256
    c(2) = (color / 256) Mod 256
    c(3) = (color / 65536) Mod 256
    ColorRGB = c
Else
    debug_err "ColorRGB", , "Color is invalid. Please check."
End If
End Function

'Turn a color to long from RGB
Public Function color(ByVal red As Long, green As Long, blue As Long) As Long
If red >= 0 And red < 256 And green >= 0 And green < 256 And blue >= 0 And blue < 256 Then
    color = blue * 65536 + green * 256 + red
Else
    debug_err "Color", , "RGB is invalid, please check."
End If
End Function

'Pause by Do events (millisecond)
Public Function Pause(ms As Long) As Long
On Error GoTo Err_Pause
    Dim S As Double, start As Double
    S = ms / 1000
    start = Timer
    Do While Timer < start + S
        DoEvents
    Loop
    Pause = ms
Exit_Pause:
    Exit Function
Err_Pause:
    debug_err "Pause", , ERR.Number & " - " & ERR.Description
    Resume Exit_Pause
End Function

'Printscreen function
Public Function AltPrintScreen(Optional ByVal actwin As Boolean) As Boolean
If actwin = True Then keybd_event VK_MENU, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
If actwin = True Then keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
AltPrintScreen = True
End Function

'Paste Picture from Clipboard
Public Function PastePicture() As IPicture
Dim bmp As Long, hand As Long, hPal As Long, lPicType As Long, Copy As Long
bmp = IsClipboardFormatAvailable(CF_BITMAP)
If bmp <> 0 Then
    'Get access to the clipboard
    OpenClipboard (0&)
    'Get a handle to the image data
    hand = GetClipboardData(CF_BITMAP)
    Copy = CopyImage(hand, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
    'Release the clipboard to other programs
    CloseClipboard
    'If we got a handle to the image, convert it into a Picture object and return it
     Set PastePicture = CreatePicture(Copy, 0, CF_BITMAP)
ElseIf bmp = 0 Then
    debug_err "PastePicture", , "No BMP images in Chipboard"
End If
End Function

'Create a picture file
Public Function CreatePicture(ByVal hPic As Long, ByVal hPal As Long, ByVal lPicType) As IPicture
    'IPicture requires a reference to "OLE Automation"
    Dim Ern As Long, uPicInfo As uPicDesc, IID_IDispatch As GUID, IPic As IPicture
    'OLE Picture types
    Const PICTYPE_BITMAP = 1
    Const PICTYPE_ENHMETAFILE = 4
    'Create the Interface GUID (for the IPicture interface)
    With IID_IDispatch
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    ' Fill uPicInfo with necessary parts.
    With uPicInfo
        .size = Len(uPicInfo) ' Length of structure.
        .Type = PICTYPE_BITMAP ' Type of Picture
        .hPic = hPic ' Handle to image.
        .hPal = hPal ' Handle to palette (if bitmap).
    End With
    ' Create the Picture object.
    Ern = OleCreatePictureIndirect(uPicInfo, IID_IDispatch, True, IPic)
    ' If an error occurred, show the description
    If Ern <> 0 Then
        debug_err "Create Picture: " & OLEError(Ern)
        Exit Function
    End If
    ' Return the new Picture object.
    Set CreatePicture = IPic
End Function

'OLE Error messages
Private Function OLEError(lErrNum As Long) As String
    'OLECreatePictureIndirect return values
    Const E_ABORT = &H80004004
    Const E_ACCESSDENIED = &H80070005
    Const E_FAIL = &H80004005
    Const E_HANDLE = &H80070006
    Const E_INVALIDARG = &H80070057
    Const E_NOINTERFACE = &H80004002
    Const E_NOTIMPL = &H80004001
    Const E_OUTOFMEMORY = &H8007000E
    Const E_POINTER = &H80004003
    Const E_UNEXPECTED = &H8000FFFF
    Const S_OK = &H0
    Select Case lErrNum
        Case E_ABORT
            fnOLEError = " Aborted"
        Case E_ACCESSDENIED
            fnOLEError = " Access Denied"
        Case E_FAIL
            fnOLEError = " General Failure"
        Case E_HANDLE
            fnOLEError = " Bad/Missing Handle"
        Case E_INVALIDARG
            fnOLEError = " Invalid Argument"
        Case E_NOINTERFACE
            fnOLEError = " No Interface"
        Case E_NOTIMPL
            fnOLEError = " Not Implemented"
        Case E_OUTOFMEMORY
            fnOLEError = " Out of Memory"
        Case E_POINTER
            fnOLEError = " Invalid Pointer"
        Case E_UNEXPECTED
            fnOLEError = " Unknown Error"
        Case S_OK
            fnOLEError = " Success!"
    End Select
End Function

'Save printscreen as image to the path
Public Sub SavePrintscreen(ByVal savePath As String, Optional ByVal actwin As Boolean)
On Error GoTo errHandler:
    AltPrintScreen (actwin)
    Do Until Chipboardex("CF_BITMAP") = True
        DoEvents
    Loop
    SavePicture PastePicture, savePath
errExit:
    Exit Sub
errHandler:
    debug_err "savePrintScreen", , "Save Picture: (" & ERR.Number & ") - " & ERR.Description
    Resume errExit
End Sub

'Check the type of chipboardex
Public Function Chipboardex(ByVal Typ As String) As Boolean
Dim t As Long
Const CF_TEXT = 1 'CF_TEXT: Text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. Use this format for ANSI text.
Const CF_BITMAP = 2 'CF_BITMAP: A handle to a bitmap (HBITMAP).
Const CF_TIFF = 6 'CF_TIFF: Tagged-image file format.
Const CF_DIB = 8 'CF_DIB: memory object containing a BITMAPINFO structure followed by the bitmap bits.
Const CF_UNICODETEXT = 13 'CF_UNICODETEXT: Unicode text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
Const CF_WAVE = 12 'CF_WAVE: Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM.
Select Case Typ
    Case "CF_TEXT"
        t = CF_TEXT
    Case "CF_BITMAP"
        t = CF_BITMAP
    Case "CF_TIFF"
        t = CF_TIFF
    Case "CF_DIB"
        t = CF_DIB
    Case "CF_UNICODETEXT"
        t = CF_UNICODETEXT
    Case "CF_WAVE"
        t = CF_WAVE
    Case Else
        t = 0
End Select
If t = 0 Then
    debug_err "Chipboardex", , "Please specify the type of the chipboard!"
    Exit Function
End If
If IsClipboardFormatAvailable(t) <> 0 Then
    Chipboardex = True
ElseIf IsClipboardFormatAvailable(t) = 0 Then
    Chipboardex = False
End If
End Function

'Excel file to ADODB Connection(Database)
Function xls_to_ADODB(ByVal dbpath As String) As ADODB.Connection
Dim sconnect As String
Dim Cn As New ADODB.Connection
'MSDASQL
'sconnect = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & DBPath & ";HDR=Yes';"
'OLEDB
sconnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath _
    & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
Cn.Open sconnect
Set xls_to_ADODB = Cn
End Function

'Folder(for csv file) to ADODB Connection(DataBase)
Public Function csv_ADODB(ByVal path As String) As ADODB.Connection
Dim Cn As New ADODB.Connection
Cn.openn ("Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & path & ";" & _
               "Extended Properties=""text; HDR=Yes; FMT=Delimited; IMEX=1;""")
RS.ActiveConnection = Cn
Set getData = Cn
End Function

'Array to recordset
Private Function ayrs(ByVal A As Variant, ByVal hdr As Boolean) As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim R As Long
Dim c As Long
Dim S As Long

Set RS = New ADODB.Recordset
If hdr = True Then
    S = 2
    For c = 1 To UBound(A, 2)
        RS.Fields.Append CStr(A(1, c)), adVariant
    Next c
Else
    S = 1
    For c = 1 To UBound(A, 2)
        RS.Fields.Append "Fld" & c, adVariant
    Next c
End If

RS.Open
For R = S To UBound(A, 1)
    RS.AddNew
    For c = 1 To UBound(A, 2)
        RS.Fields(c - 1).Value = A(R, c)
    Next c
    RS.Update
Next R
RS.Filter = ""
Set ayrs = RS
End Function










