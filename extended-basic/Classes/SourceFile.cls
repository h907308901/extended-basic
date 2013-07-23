VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SourceFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_FileName As String
Dim m_FileHandle As Long
Dim m_FileSize As Long
Dim m_FileMappingHandle As Long
Dim m_Data As Long 'point to data buffer
Dim m_Pointer As Long 'pointer in buffer
Dim m_Lines() As Long '1-based
Dim m_CurrentLine As Long
Dim m_Index As Long

Dim m_SymTable As Dictionary
Dim m_hModule As Long

Friend Function OpenSource(FileName As String, ByVal nIndex As Long) As Long
    Dim l As LARGE_INTEGER, b As Byte
    Dim Line As Long, i As Long
    Dim s As String, i2 As Long
    If m_FileHandle <> 0 Then
        FatalError "SourceFile 0x" & Hex(ObjPtr(Me)) & " has already opened file '" & m_FileName & "'"
    End If
    m_Index = nIndex
    'open file
    m_FileName = FileName
    m_FileHandle = CreateFile(FileName, FILE_ALL_ACCESS, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_EXISTING, 0, 0)
    If m_FileHandle = -1 Then
        GeneralError "cannot open file '" & FileName & "'"
        GoTo ErrLine
    End If
    'get size
    GetFileSizeEx m_FileHandle, l
    If l.HighPart <> 0 Or l.LowPart > &H1000000 Then 'max size 16M
        GeneralError "file '" & FileName & "' is too big"
        GoTo ErrLine
    End If
    m_FileSize = l.LowPart
    'create file mapping
    m_FileMappingHandle = CreateFileMapping(m_FileHandle, ByVal 0, PAGE_READONLY, 0, 0, vbNullString)
    If m_FileMappingHandle = 0 Then
        GeneralError "cannot create file mapping for file '" & FileName & "'"
        GoTo ErrLine
    End If
    'map
    m_Data = MapViewOfFile(m_FileMappingHandle, FILE_MAP_READ, 0, 0, 0)
    If m_Data = 0 Then
        GeneralError "cannot map file '" & FileName & "'"
        GoTo ErrLine
    End If
    'calculate lines
    'Note: 0th and last of m_Lines are for fault tolerance
    m_Pointer = 0
    Line = 1
    ReDim m_Lines(0 To Line) As Long
    m_Lines(0) = &H80000000
    For i = 0 To m_FileSize
        GetMem1 m_Data + i, b
        If b = 13 Then
            Line = Line + 1
            ReDim Preserve m_Lines(0 To Line) As Long
            m_Lines(Line) = i
        ElseIf b > &H7F Then 'double-byte
            i = i + 1
        End If
    Next
    Line = Line + 1
    ReDim Preserve m_Lines(0 To Line) As Long
    m_Lines(Line) = &H7FFFFFFF
    m_CurrentLine = 1
    'initialize symbol table
    Set m_SymTable = New Dictionary
    'create module
    i = InStrRev(FileName, "\")
    i2 = InStrRev(FileName, "/")
    If i < i2 Then i = i2
    i = Len(FileName) - i
    s = Right$(FileName, i)
    i = InStr(s, ".") - 1
    If i > 0 Then s = Left$(s, i)
    m_hModule = LLVMModuleCreateWithName(StrPtrA(s))
    If m_hModule = 0 Then
        GeneralError "cannot create module '" & s & "' for file '" & FileName & "'"
        GoTo ErrLine
    End If
    OpenSource = 0
    Exit Function
ErrLine:
    OpenSource = -1
    CloseSource
End Function

Friend Sub CloseSource()
    If m_FileHandle = 0 Then Exit Sub
    Set m_SymTable = Nothing
    LLVMDisposeModule m_hModule
    UnmapViewOfFile m_Data
    CloseHandle m_FileMappingHandle
    CloseHandle m_FileHandle
    m_FileHandle = 0
End Sub

Friend Function Reverse() As String
    Dim v As Variant, Node As IASTNode
    For Each v In m_SymTable
        Set Node = m_SymTable(v)
        Reverse = Reverse & Node.Reverse(0) & vbCrLf
    Next
End Function

Friend Function Codegen(ByVal C As Context) As Long
    Dim v As Variant, Node As IASTNode
    C.EnterFile Me
    For Each v In m_SymTable
        Set Node = m_SymTable(v)
        Node.Codegen C
    Next
    C.ExitFile
    Codegen = 1
End Function

Friend Property Get Index() As Long
    Index = m_Index
End Property

Friend Property Get SymTable() As Dictionary
    Set SymTable = m_SymTable
End Property

Friend Property Get hModule() As Long
    hModule = m_hModule
End Property

Friend Function GetFileName() As String
    GetFileName = m_FileName
End Function

Friend Function GetChar() As VBCharType
    If m_Pointer > m_FileSize Then
        GetChar = [EOF]
        Exit Function
    End If
    GetMem1 m_Data + m_Pointer, GetChar
    m_Pointer = m_Pointer + 1
    If GetChar > &H7F& Then 'double-byte
        GetChar = GetChar + GetChar() * &H100& '???
    End If
End Function

Friend Function GetLine() As Long ''' TODO:
    Do
        If m_Pointer > m_Lines(m_CurrentLine) Then
            If m_Pointer <= m_Lines(m_CurrentLine + 1) Then
                Exit Do
            Else
                m_CurrentLine = m_CurrentLine + 1
            End If
        Else
            m_CurrentLine = m_CurrentLine - 1
        End If
    Loop
    GetLine = m_CurrentLine
End Function

Friend Sub UnGetChar(Optional ByVal uNumber = 1)
    m_Pointer = m_Pointer - uNumber
End Sub

Private Sub Class_Terminate()
    CloseSource
End Sub
