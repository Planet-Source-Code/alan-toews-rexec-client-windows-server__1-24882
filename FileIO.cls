VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'I've written most of this class, but the
'huffman compression/decmpression algorithm was taken
'verbatim from James Vincent Carnicelli's project.



Private Enum HuffmanTreeNodeParts
    htnWeight = 1
    htnIsLeaf = 2
    htnAsciiCode = 3
    htnBitCode = 4
    htnLeftSubtree = 5
    htnRightSubtree = 6
End Enum

Private Type BrowseInfo
    hwndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
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
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Public Enum CSIDLS
    CSIDL_DESKTOP = &H0
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_ALL_USERS_STARTMENU = &H16
    CSIDL_ALL_USERS_PROGRAMS = &H17
    CSIDL_ALL_USERS_SATRTUP = &H18
    CSIDL_ALL_USERS_DESKTOP = &H19
    CSIDL_APPLICATION_DATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPLICATION_DATA = &H1C
    CSIDL_ALL_USERS_FAVOURITES = &H1F
    CSIDL_TEMPORARY_INTERNET_FILES = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_ALL_USERS_APPLICATION_DATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM_32 = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_PICTURES = &H27
    CSIDL_USER_DIRECTORY = &H28
    CSIDL_COMMON_FILES = &H2B
    CSIDL_ALL_USERS_TEMPLATES = &H2D
    CSIDL_ALL_USERS_DOCUMENTS = &H2E
    CSIDL_ALL_USERS_ADMINISTRATIVE_TOOLS = &H2
End Enum
Private Const MAX_PATH = 260


Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private bClosing As Boolean

Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (bBrowse As BrowseInfo) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    
Private Const PROCESS_ALL_ACCESS& = &H1F0FFF
Private Const INFINITE = &HFFFF
Private Const STILL_ACTIVE = &H103&

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31


Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum

Public Function RunProgram(sDocName As String, _
                    Optional ByVal Action As String = "Open", _
                    Optional ByVal Parameters As String = vbNullString, _
                    Optional ByVal Directory As String = vbNullString, _
                    Optional ByVal WindowState As StartWindowState = START_NORMAL) As Long
    Dim Response As Long
    Response = ShellExecute(&O0, Action, sDocName, Parameters, Directory, WindowState)
    Select Case Response
        Case Is < 33
            RunProgram = 0
        Case Else
            RunProgram = Response
    End Select
End Function
Public Function RunAndWait(sDocName As String, _
                    Optional ByVal Action As String = "Open", _
                    Optional ByVal Parameters As String = vbNullString, _
                    Optional ByVal Directory As String = vbNullString, _
                    Optional ByVal WindowState As StartWindowState = START_NORMAL) As Boolean
    Dim lReturn As Long, lWaitReturn As Long
    
    If Directory <> vbNullString Then If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    
    lReturn = Shell(Directory & sDocName & " " & Parameters, WindowState)
    lWaitReturn = WaitForProgram(lReturn)
    If lReturn = 0 Then
        RunAndWait = False
    Else
        RunAndWait = True
    End If
End Function

' Save values into the file.
Public Function Save(ByVal TextToSave As Variant, ByVal file_name As Variant) As Boolean
    On Error GoTo ErrorCode
    
    Dim fNum As Integer
    Dim I As Integer
    Dim txt As String

    fNum = FreeFile
        
    Open file_name For Output As fNum   ' Open the file.
    Print #fNum, TextToSave             ' Write the text
    Close fNum                          ' Close the file.
    
    Save = True
    Exit Function
ErrorCode:
    Save = False
End Function

Public Function Load(ByVal file_name As Variant) As String
    Dim fNum As Integer, x As Integer
    Dim inText(32766) As Byte, file_text As String
    On Error GoTo ErrorCode
    fNum = FreeFile
    
    Open file_name For Random As fNum Len = UBound(inText) + 1
        Do While Not EOF(fNum) ' Loop until end of file.
            Get #fNum, , inText
            For x = LBound(inText) To UBound(inText)
                If inText(x) = 0 Then Exit For
                file_text = file_text & Chr(inText(x))
            Next
        Loop
    Close fNum
    Load = file_text
    
    Exit Function
ErrorCode:
    If Err.Number = 53 Then
        Load = "FILE NOT FOUND"
        Close fNum
        Exit Function
    Else
        MsgBox Err.Description & vbCrLf & "Error " & Err.Number, vbExclamation, Err.Source & " Error"
        Close fNum
    End If
End Function
Public Function FileLength(ByVal Filename As Variant) As Long
Dim length As Long

    On Error GoTo FileDoesntExist
    length = FileLen(Filename)
    FileLength = length
    Exit Function
    
FileDoesntExist:
    FileLength = -1
End Function



Public Function Delete(Filename As Variant) As Boolean
    On Error GoTo DelError
    Kill Filename
    Delete = True
    Exit Function
DelError:
    Delete = False
End Function


Public Function Copy(Source As Variant, Destination As Variant) As Boolean
    On Error GoTo CopyError
    FileCopy Source, Destination
    Copy = True
    Exit Function
CopyError:
    Copy = False
    
End Function

Public Function Rename(FromName As Variant, ToName As Variant) As Boolean
    On Error GoTo MoveError
    
    Name FromName As ToName
    Rename = True
    Exit Function
MoveError:
    Rename = False
    
End Function
Public Function Move(Source As Variant, Destination As Variant) As Boolean
    On Error GoTo MoveError
    
    FileCopy Source, Destination
    Kill Source
    Move = True
    Exit Function
MoveError:
    Move = False
    
End Function


Public Function MakeDIR(path As Variant) As Boolean
    On Error GoTo DIRError
    MkDir path
    MakeDIR = True
    Exit Function
DIRError:
    MakeDIR = False
    
End Function


Public Function RemoveDIR(path As Variant) As Boolean
    On Error GoTo DIRError2
    RmDir path
    RemoveDIR = True
    Exit Function
DIRError2:
    RemoveDIR = False
    
End Function
Public Function Exists(Filename) As Boolean
    On Error GoTo ErrorCode
    If FileLen(Filename) >= 0 Then Exists = True
    
    Exit Function
ErrorCode:
    Exists = False
    
End Function

Private Function WaitForProgram(ByVal lProgram As Long) As Long
    Dim lProgramCode As Long, hProgram As Long
    
    'get the rpogram handle
    hProgram = OpenProcess(PROCESS_ALL_ACCESS, False, lProgram)
    'check if the program has exited
    GetExitCodeProcess hProgram, lProgramCode

    'wait while the program is still runnung
    Do While (lProgramCode = STILL_ACTIVE) And Not bClosing
        'check if the program has exited.
        GetExitCodeProcess hProgram, lProgramCode
        DoEvents

    Loop
    'clean up
    CloseHandle hProgram
    WaitForProgram = lProgramCode
    End Function


Public Function GetSpecialfolder(CSIDL As CSIDLS) As String
    Dim r As Long, path As String
    Dim IDL As ITEMIDLIST
    'Get the special folder
    r = SHGetSpecialFolderLocation(0&, CSIDL, IDL)
    If r = 0 Then
        'Create a buffer
        path = Space(512)
        'Get the path from the IDList
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal path)
        'Remove the unnecessary nulls
        GetSpecialfolder = Left(path, InStr(path, vbNullChar) - 1)
    Else
        GetSpecialfolder = ""
    End If
End Function

Public Function Obscure(txt As String) As String
    Dim x As Long, result As String
    For x = 1 To Len(txt)
        result = result & Format(123 + Asc(Mid(txt, x, 1)), "000")
    Next
    Obscure = result
    
End Function
Public Function Recover(ObscureTxt As String) As String
    Dim x As Long, result As String
    For x = 1 To Len(ObscureTxt) Step 3
        result = result & Chr(Mid(ObscureTxt, x, 3) - 123)
    Next
    Recover = result
End Function
'################################################################
' Huffman Coding Compression / Decompression Algorithm




'Compress the text.
Public Function Compress(Text As String, Optional Force As Boolean) As String
    Dim TextLen As Long, Char As Byte, I As Long, j As Long
    Dim CodeCounts(255) As Long, BitStrings(255), BitString
    Dim HuffmanTrees As Collection
    Dim HTRootNode As Collection, HTNode As Collection
    Dim NextByte As Byte, BitPos As Integer, Temp As String
    
    'Initialize for processing.
    TextLen = Len(Text)
    Set HuffmanTrees = New Collection
    
    'Is there anything to encode?
    If TextLen = 0 Then
        Compress = "HE0" & vbCr  'Version 0 = Plain text
        Exit Function  'No point in continuing
    End If
    
    Compress = "HE2" & vbCr  'Version 1
    
    'Count how many times each ASCII code is encountered in text.
    For I = 1 To TextLen
        Char = Asc(Mid(Text, I, 1))
        CodeCounts(Char) = CodeCounts(Char) + 1
    Next
    
    'Initialize the forest of Huffman trees; one for each ASCII
    'character used.
    For I = 0 To UBound(CodeCounts)
        If CodeCounts(I) > 0 Then
            Set HTNode = NewNode
            S HTNode, htnAsciiCode, Chr(I)
            S HTNode, htnWeight, CDbl(CodeCounts(I) / TextLen)
            S HTNode, htnIsLeaf, True
            
            'Now place it in its reverse-ordered position.
            For j = 1 To HuffmanTrees.Count + 1
                If j > HuffmanTrees.Count Then
                    HuffmanTrees.Add HTNode
                    Exit For
                End If
                If HTNode(htnWeight) >= HuffmanTrees(j)(htnWeight) Then
                    HuffmanTrees.Add HTNode, , j
                    Exit For
                End If
            Next
        End If
    Next
    
    'Now assemble all these single-level Huffman trees into
    'one single tree, where all the leaves have the ASCII codes
    'associated with them.
    If HuffmanTrees.Count = 1 Then
        Set HTNode = NewNode
        S HTNode, htnLeftSubtree, HuffmanTrees(1)
        S HTNode, htnWeight, 1
        HuffmanTrees.Remove (1)
        HuffmanTrees.Add HTNode
    End If
    While HuffmanTrees.Count > 1
        Set HTNode = NewNode
        S HTNode, htnRightSubtree, HuffmanTrees(HuffmanTrees.Count)
        HuffmanTrees.Remove HuffmanTrees.Count
        S HTNode, htnLeftSubtree, HuffmanTrees(HuffmanTrees.Count)
        HuffmanTrees.Remove HuffmanTrees.Count
        S HTNode, htnWeight, HTNode(htnLeftSubtree)(htnWeight) + HTNode(htnRightSubtree)(htnWeight)
        
        'Place this new tree it in its reverse-ordered position.
        For j = 1 To HuffmanTrees.Count + 1
            If j > HuffmanTrees.Count Then
                HuffmanTrees.Add HTNode
                Exit For
            End If
            If HTNode(htnWeight) >= HuffmanTrees(j)(htnWeight) Then
                HuffmanTrees.Add HTNode, , j
                Exit For
            End If
        Next
    Wend
    Set HTRootNode = HuffmanTrees(1)
    AttachBitCodes BitStrings, HTRootNode, Array()
    For I = 0 To UBound(BitStrings)
        If Not IsEmpty(BitStrings(I)) Then
            Set HTNode = BitStrings(I)
            Temp = Temp & HTNode(htnAsciiCode) _
              & BitsToString(HTNode(htnBitCode))
        End If
    Next
    Compress = Compress & Len(Temp) & vbCr & Temp
    
    'The next part of the header is a checksum value, which
    'we'll use later to verify our decompression.
    Char = 0
    For I = 1 To TextLen
        Char = Char Xor Asc(Mid(Text, I, 1))
    Next
    Compress = Compress & Chr(Char)
    
    'The final part of the header identifies how many bytes
    'the original text strings contains.  We will probably
    'have a few unused bits in the last byte that we need to
    'account for.  Additionally, this serves as a final check
    'for corruption.
    Compress = Compress & TextLen & vbCr
    
    'Now we can encode the data by exchanging each ASCII byte for
    'its appropriate bit string.
    BitPos = -1
    Char = 0
    Temp = ""
    For I = 1 To TextLen
        BitString = BitStrings(Asc(Mid(Text, I, 1)))(htnBitCode)
        'Add each bit to the end of the output stream's 1-byte buffer.
        For j = 0 To UBound(BitString)
            BitPos = BitPos + 1
            If BitString(j) = 1 Then
                Char = Char + 2 ^ BitPos
            End If
            'If the bit buffer is full, dump it to the output stream.
            If BitPos >= 7 Then
                Temp = Temp & Chr(Char)
                'If the temporary output buffer is full, dump it
                'to the final output stream.
                If Len(Temp) > 1024 Then
                    Compress = Compress & Temp
                    Temp = ""
                End If
                BitPos = -1
                Char = 0
            End If
        Next
    Next
    If BitPos > -1 Then
        Temp = Temp & Chr(Char)
    End If
    If Len(Temp) > 0 Then
        Compress = Compress & Temp
    End If
    
    'If it takes up more space compressed because the source is
    'small and the header is big, we'll leave it uncompressed
    'and prepend it with a 4 byte header.
    If Len(Compress) > TextLen And Not Force Then
        Compress = "HE0" & vbCr & Text
    End If
End Function


'Decompress the string back into its original text.
Public Function Decompress(ByVal Text As String) As String
    Dim Pos As Long, Temp As String, Char As Byte, Bits
    Dim I As Long, j As Long, CharsFound As Long, BitPos As Integer
    Dim CheckSum As Byte, SourceLen As Long, TextLen As Long
    Dim HTRootNode As Collection, HTNode As Collection
    
    'If this was left uncompressed, this will be easy.
    If Left(Text, 4) = "HE0" & vbCr Then
        Decompress = Mid(Text, 5)
        Exit Function
    End If
    
    'If this is any version other than 2, we'll bow out.
    If Left(Text, 4) <> "HE2" & vbCr Then
        Err.Raise vbObjectError, "Decompress()", _
          "The data either was not compressed with HE2 or is corrupt"
    End If
    Text = Mid(Text, 5)
    
    'Extract the ASCII character bit-code table's byte length.
    Pos = InStr(1, Text, vbCr)
    If Pos = 0 Then
        Err.Raise vbObjectError, "Decompress()", _
          "The data either was not compressed with HE2 or is corrupt"
    End If
    On Error Resume Next
    TextLen = Left(Text, Pos - 1)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError, "Decompress()", _
          "The header is corrupt"
    End If
    On Error GoTo 0
    Text = Mid(Text, Pos + 1)
    Temp = Left(Text, TextLen)
    Text = Mid(Text, TextLen + 1)
    'Now extract the ASCII character bit-code table.
    Set HTRootNode = NewNode
    Pos = 1
    While Pos <= Len(Temp)
        Char = Asc(Mid(Temp, Pos, 1))
        Pos = Pos + 1
        Bits = StringToBits(Pos, Temp)
        Set HTNode = HTRootNode
        For j = 0 To UBound(Bits)
            If Bits(j) = 1 Then
                If HTNode(htnLeftSubtree) Is Nothing Then
                    S HTNode, htnLeftSubtree, NewNode
                End If
                Set HTNode = HTNode(htnLeftSubtree)
            Else
                If HTNode(htnRightSubtree) Is Nothing Then
                    S HTNode, htnRightSubtree, NewNode
                End If
                Set HTNode = HTNode(htnRightSubtree)
            End If
        Next
        S HTNode, htnIsLeaf, True
        S HTNode, htnAsciiCode, Chr(Char)
        S HTNode, htnBitCode, Bits
    Wend
    
    'Extract the checksum.
    CheckSum = Asc(Left(Text, 1))
    Text = Mid(Text, 2)
    
    'Extract the length of the original string.
    Pos = InStr(1, Text, vbCr)
    If Pos = 0 Then
        Err.Raise vbObjectError, "Decompress()", _
          "The header is corrupt"
    End If
    On Error Resume Next
    SourceLen = Left(Text, Pos - 1)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError, "Decompress()", _
          "The header is corrupt"
    End If
    On Error GoTo 0
    Text = Mid(Text, Pos + 1)
    TextLen = Len(Text)
    
    'Now that we've processed the header, let's decode the actual data.
    I = 1
    BitPos = -1
    Set HTNode = HTRootNode
    Temp = ""
    While CharsFound < SourceLen
        If BitPos = -1 Then
            If I > TextLen Then
                Err.Raise vbObjectError, "Decompress()", _
                  "Expecting more bytes in data stream"
            End If
            Char = Asc(Mid(Text, I, 1))
            I = I + 1
        End If
        BitPos = BitPos + 1
        
        If (Char And 2 ^ BitPos) > 0 Then
            Set HTNode = HTNode(htnLeftSubtree)
        Else
            Set HTNode = HTNode(htnRightSubtree)
        End If
        If HTNode Is Nothing Then
            'Uh oh.  We've followed the tree to a Huffman tree to a dead
            'end, which won't happen unless the data is corrupt.
            Err.Raise vbObjectError, "Decompress()", _
              "The header (lookup table) is corrupt"
        End If
        
        If HTNode(htnIsLeaf) Then
            Temp = Temp & HTNode(htnAsciiCode)
            If Len(Temp) > 1024 Then
                Decompress = Decompress & Temp
                Temp = ""
            End If
            CharsFound = CharsFound + 1
            Set HTNode = HTRootNode
        End If
        
        If BitPos >= 7 Then BitPos = -1
    Wend
    If Len(Temp) > 0 Then
        Decompress = Decompress & Temp
    End If
    If I <= TextLen Then
        Err.Raise vbObjectError, "Decompress()", _
          "Found extra bytes at end of data stream"
    End If
    
    'Verify data to check for corruption.
    If Len(Decompress) <> SourceLen Then
        Err.Raise vbObjectError, "Decompress()", _
          "Data corrupt because check sums do not match"
    End If
    Char = 0
    For I = 1 To SourceLen
        Char = Char Xor Asc(Mid(Decompress, I, 1))
    Next
    If Char <> CheckSum Then
        Err.Raise vbObjectError, "Decompress()", _
          "Data corrupt because check sums do not match"
    End If
End Function



'----------------------------------------------------------------
' Everything below here is only for supporting the two main
' routines above.
'----------------------------------------------------------------


'Follows the tree, now built, to its end leaf nodes, where the
'character codes are, in order to tell those character codes
'what their bit string representations are.
Private Sub AttachBitCodes(BitStrings, HTNode As Collection, ByVal Bits)
    If HTNode Is Nothing Then Exit Sub
    If HTNode(htnIsLeaf) Then
        S HTNode, htnBitCode, Bits
        Set BitStrings(Asc(HTNode(htnAsciiCode))) = HTNode
    Else
        ReDim Preserve Bits(UBound(Bits) + 1)
        Bits(UBound(Bits)) = 1
        AttachBitCodes BitStrings, HTNode(htnLeftSubtree), Bits
        Bits(UBound(Bits)) = 0
        AttachBitCodes BitStrings, HTNode(htnRightSubtree), Bits
    End If
End Sub

'Turns a string of '0' and '1' characters into a string of bytes
'containing the bits, preceeded by 1 byte indicating the
'number of bits represented.
Private Function BitsToString(Bits) As String
    Dim Char As Byte, I As Long
    BitsToString = Chr(UBound(Bits) + 1)  'Number of bits
    For I = 0 To UBound(Bits)
        If I Mod 8 = 0 Then
            If I > 0 Then BitsToString = BitsToString & Chr(Char)
            Char = 0
        End If
        If Bits(I) = 1 Then  'Bit value = 1
            'Mask the bit into its proper position in the byte
            Char = Char + 2 ^ (I Mod 8)
        End If
    Next
    BitsToString = BitsToString & Chr(Char)
End Function

'The opposite of BitsToString() function.
Private Function StringToBits(StartPos As Long, Bytes As String)
    Dim Char As Byte, I As Long, BitCount As Long, Bits
    Bits = Array()
    BitCount = Asc(Mid(Bytes, StartPos, 1))
    StartPos = StartPos + 1
    For I = 0 To BitCount - 1
        If I Mod 8 = 0 Then
            Char = Asc(Mid(Bytes, StartPos, 1))
            StartPos = StartPos + 1
        End If
        ReDim Preserve Bits(UBound(Bits) + 1)
        If (Char And 2 ^ (I Mod 8)) > 0 Then   'Bit value = 1
            Bits(UBound(Bits)) = 1
        Else  'Bit value = 0
            Bits(UBound(Bits)) = 0
        End If
    Next
    StringToBits = Bits
End Function

'Remove the specified item and put the specified value in its place.
Private Sub S(Col As Collection, Index As HuffmanTreeNodeParts, Value)
    Col.Remove Index
    If Index > Col.Count Then
        Col.Add Value
    Else
        Col.Add Value, , Index
    End If
End Sub

'Creates a new Huffman tree node with the default values set.
Private Function NewNode() As Collection
    Dim Node As New Collection
    Node.Add 0  'htnWeight
    Node.Add False  'htnIsLeaf
    Node.Add Chr(0)  'htnAsciiCode
    Node.Add ""  'htnBitCode
    Node.Add Nothing  'htnLeftSubtree
    Node.Add Nothing  'htnRightSubtree
    Set NewNode = Node
End Function




Private Sub Class_Initialize()
    bClosing = False
End Sub

Private Sub Class_Terminate()
    bClosing = True
End Sub
