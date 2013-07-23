Attribute VB_Name = "mdlMain"
Option Explicit

'****************************************
' Main
' main function, command parser, standard output support
'****************************************

''' Note: if RELEASE_VERSION is set, VB6 default err handling and IDE debugging will be invalid
#Const RELEASE_VERSION = False

#If Not RELEASE_VERSION Then
''' WARNING: if ENABLE_IDE_DEBUG is not set, do not run in IDE mode
''' Note: if there is error break during running in IDE, please restart VB6
#Const ENABLE_IDE_DEBUG = True
'dump module when verification failed
#Const DEBUG_DUMP_MODULE = True
'dump reversed code
#Const DEBUG_DUMP_REVERSE = False

#End If

Public Const VERSION = "prealpha version"

Public Const DEFAULT_TRIPLE = "i686-pc-mingw32"
Public Const DEFAULT_FEATURES = "i686,mmx,cmov,sse,sse2,sse3"

#If ENABLE_IDE_DEBUG Then
Public g_InIDE As Boolean
#End If

Dim hStdOut As Long
Dim WarningCount As Long, ErrorCount As Long
' === Commands ===
Public g_FileCount As Long
Public g_FileList() As String '1-based
Public g_OutputFile() As String '1-based
Public g_OutputType As Long 'output file type, 0 means no output
Public Const OT_BITCODE = 1 '-BC
Public Const OT_LLVMFILE = 2 '-LL
Public Const OT_ASSEMBLY = 3 '-AS
Public Const OT_OBJECT = 4 '-OB
Public g_OptLevel 'optimization level
Public g_DefaultCC As Long 'default call conv
Public g_h As Boolean 'help -?
Public g_Vd As Boolean 'verification disabled -Vd
' === Commands ===
Public g_Triple As String
Public g_Features As String
Public g_TargetMachine As Long
Public g_TargetData As Long
Public g_TargetDataLayout(2047) As Byte
Dim Files As Dictionary

Sub Main()
    Dim i As Long, s As String, Argc As Long, Argv() As String
#If RELEASE_VERSION Then
    On Error GoTo ErrLine
#End If
#If ENABLE_IDE_DEBUG Then
        g_InIDE = App.LogMode <> 1
        If g_InIDE Then
            s = "eb test.eb -O0 -LL"
        Else
#End If
            s = App.EXEName & " " & Command$
#If ENABLE_IDE_DEBUG Then
        End If
#End If
    'get commands
    Do
        If s = vbNullString Then Exit Do
        If Left$(s, 1) = """" Then
            i = InStr(2, s, """")
            If i = 0 Then
                Argc = Argc + 1
                ReDim Preserve Argv(Argc - 1)
                Argv(Argc - 1) = Mid$(s, 2)
                Exit Do
            Else
                Argc = Argc + 1
                ReDim Preserve Argv(Argc - 1)
                Argv(Argc - 1) = Mid$(s, 2, i - 2)
                s = Mid$(s, i + 1)
            End If
        Else
            i = InStr(1, s, " ")
            If i = 0 Then
                Argc = Argc + 1
                ReDim Preserve Argv(Argc - 1)
                Argv(Argc - 1) = s
                Exit Do
            Else
                Argc = Argc + 1
                ReDim Preserve Argv(Argc - 1)
                Argv(Argc - 1) = Left$(s, i - 1)
                s = Mid$(s, i + 1)
            End If
        End If
    Loop
    'actual compilation step
#If ENABLE_IDE_DEBUG Then
    If g_InIDE Then
        Debug.Print "CompilerMain returned with code 0x" & Hex$(CompilerMain(Argc, Argv))
        End
    Else
#End If
        ExitProcess CompilerMain(Argc, Argv)
#If ENABLE_IDE_DEBUG Then
    End If
#End If
#If RELEASE_VERSION Then
ErrLine:
    FatalAppExit -1, "internal error #" & Err.Number & ": " & Err.Description
#End If
End Sub

Function CompilerMain(ByVal Argc As Long, Argv() As String) As Long
    Dim Src As SourceFile
    Dim Node As IASTNode
    Dim C As New Context
    Dim outmode As Long, hStream As Long, hRawStream As Long
    Dim hPass As Long, hFunction As Long
    Dim Threshold As Long
    Dim v As Variant, i As Long, s As String, l As Long, Ret As Long
    Dim OutMsg As Long
    'compiler initialization step
    'initialize standard output
    ConsoleInit
    'set default values
    CommandDefaultValueSet
    'parse commands
    Ret = CommandParser(Argc, Argv)
    If g_h Then
        ShowHelp
        GoTo Final
    End If
    If Ret = -1 Then GoTo ErrLine
    'check output file name
    If g_FileCount = 0 Then
        GeneralError "no input file"
        GoTo ErrLine
    End If
    For i = 1 To g_FileCount
        If g_OutputFile(i) = vbNullString Then
            l = InStrRev(g_FileList(1), ".") - 1
            If l = 0 Then 'no dot
                s = g_FileList(1)
            ElseIf l < InStrRev(g_FileList(1), "\") Or l < InStrRev(g_FileList(1), "/") Then ' dot not in file name
                s = g_FileList(1)
            Else
                s = Left$(g_FileList(1), l)
            End If
            Select Case g_OutputType
            Case OT_BITCODE
                s = s & ".bc"
            Case OT_LLVMFILE
                s = s & ".ll"
            Case OT_ASSEMBLY
                s = s & ".asm"
            Case OT_OBJECT
                s = s & ".obj"
            End Select
            g_OutputFile(i) = s
        End If
    Next
    'initialize llvm
    Ret = LLVMInit
    If Ret = -1 Then GoTo ErrLine
    'open source files
    Set Files = New Dictionary
    For i = 1 To g_FileCount
        Set Src = New SourceFile
        Ret = Src.OpenSource(g_FileList(i), i)
        If Ret Then GoTo ErrLine
        Files.Add Src.GetFileName, Src
    Next
    Select Case g_OutputType
    Case OT_LLVMFILE
        outmode = LLVMOpenMode_out
    Case OT_ASSEMBLY
        outmode = LLVMOpenMode_out Or LLVMOpenMode_binary
    Case OT_OBJECT
        outmode = LLVMOpenMode_binary
    Case Else
        outmode = 0
    End Select
    For Each v In Files
        Set Src = Files(v)
        'initialize lexer
        LexerInit
        'initialize syntax
        SyntaxInit
        'initialize type
        TypeInit
        'initialize runtime
        RuntimeInit
        'parse files
        PrintLine "Parsing file '" & Src.GetFileName & "' ..."
        Ret = ParseFile(Src)
        If ErrorCount > 0 Then GoTo CompileErr
#If DEBUG_DUMP_REVERSE Then
        PrintLine Src.Reverse
#End If
        'generate code
        PrintLine "Generating code for '" & Src.GetFileName & "' ..."
        For i = 1 To 29
            C.Step = 2 ^ i
            If C.Step And cg_valid_mask Then
                Src.Codegen C
            End If
        Next
        If Not g_Vd Then
            Ret = LLVMVerifyModule(Src.hModule, LLVMPrintMessageAction, OutMsg)
            If Ret Then
#If DEBUG_DUMP_MODULE Then
                LLVMDumpModule Src.hModule
#End If
                LLVMDisposeMessage OutMsg
                PrintError Src.GetFileName, 0, "verification failed, please contect developers"
            End If
        End If
#If ENABLE_IDE_DEBUG Then
        PrintOutputFile
#End If
        If ErrorCount > 0 Then GoTo CompileErr
        'optimization
        hPass = LLVMCreateFunctionPassManagerForModule(Src.hModule)
        LLVMAddTargetData g_TargetData, hPass
        LLVMCreateStandardFunctionPasses hPass, g_OptLevel
        LLVMInitializeFunctionPassManager hPass
        hFunction = LLVMGetFirstFunction(Src.hModule)
        Do Until hFunction = 0
            If LLVMCountBasicBlocks(hFunction) > 0 Then
                LLVMRunFunctionPassManager hPass, hFunction
            End If
            hFunction = LLVMGetNextFunction(hFunction)
        Loop
        LLVMFinalizeFunctionPassManager hPass
        LLVMDisposePassManager hPass
        Select Case g_OptLevel
        Case LLVMCodeGenOpt_None
            Threshold = 0
        Case LLVMCodeGenOpt_Less
            Threshold = 200
        Case Else
            Threshold = 250
        hPass = LLVMCreatePassManager
        LLVMAddTargetData g_TargetData, hPass
        LLVMCreateStandardModulePasses hPass, g_OptLevel, 0, 1, 0, g_OptLevel >= LLVMCodeGenOpt_Default And 1, 1, Threshold
        'output
        End Select
        If g_OutputType <> 0 Then
            If outmode <> 0 Then
                hStream = Util_CreateOStreamFromFile(g_OutputFile(Src.Index), outmode)
                hRawStream = LLVMCreateRaw_OS_OStream(hStream)
                hRawStream = LLVMCreateFormattedRawOStream(hRawStream, 1)
                If hRawStream = 0 Then
                    GeneralError "error creating output file " & g_OutputFile(Src.Index)
                    GoTo Final
                End If
            End If
            Select Case g_OutputType
            Case OT_BITCODE
                LLVMRunPassManager hPass, Src.hModule
                LLVMWriteBitcodeToFile Src.hModule, StrPtrA(g_OutputFile(Src.Index))
            Case OT_LLVMFILE
                LLVMAddPrintModulePass hPass, hRawStream, 0, vbNullChar
                LLVMRunPassManager hPass, Src.hModule
            Case OT_ASSEMBLY
                Ret = LLVMTargetMachineAddPassesToEmitFile(g_TargetMachine, hPass, hRawStream, CGFT_AssemblyFile, g_OptLevel, 0)
                If Ret <> 0 Then
                    GeneralError "error emitting file " & g_OutputFile(Src.Index)
                    GoTo Final
                End If
                LLVMRunPassManager hPass, Src.hModule
            Case OT_OBJECT
                hPass = LLVMCreatePassManager
                LLVMAddTargetData g_TargetData, hPass
                Ret = LLVMTargetMachineAddPassesToEmitFile(g_TargetMachine, hPass, hRawStream, CGFT_ObjectFile, g_OptLevel, 0)
                If Ret <> 0 Then
                    GeneralError "error emitting file " & g_OutputFile(Src.Index)
                    GoTo Final
                End If
                LLVMRunPassManager hPass, Src.hModule
            End Select
        End If
        If outmode <> 0 Then
            LLVMDisposeRaw_OStream hRawStream
            Util_DisposeOStream hStream
        End If
        LLVMDisposePassManager hPass
    Next
    GoTo CompileOk
CompileErr:
    Src.CloseSource
CompileOk:
    PrintCounts
    GoTo Final
ErrLine:
    PrintLine "use ""-?"" for help"
Final:
    Set Files = Nothing
    RuntimeExit
    TypeExit
    LLVMExit
    ConsoleExit
End Function

Public Function LLVMInit() As Long
    Dim i As Long
    LLVMInitializeAllTargetInfos
    LLVMInitializeAllTargets
    LLVMInitializeAllAsmPrinters
    LLVMInitializeAllAsmParsers
    g_TargetMachine = LLVMCreateTargetMachine(g_Triple, g_Features)
    If g_TargetMachine = 0 Then
        GeneralError "cannot create target machine for triple '" & g_Triple & "' and features '" & g_Features & "'"
        GoTo ErrLine
    End If
    i = VarPtr(g_TargetDataLayout(0))
    LLVMTargetMachineGetDataLayout g_TargetMachine, ByVal i, 2048
    g_TargetData = LLVMCreateTargetData(i)
    Exit Function
ErrLine:
    LLVMInit = -1
End Function

Public Sub LLVMExit()
    LLVMDisposeTargetData g_TargetData
    LLVMDisposeTargetMachine g_TargetMachine
End Sub

Public Sub CommandDefaultValueSet()
    g_Triple = DEFAULT_TRIPLE
    g_Features = DEFAULT_FEATURES & vbNullChar
    g_OptLevel = LLVMCodeGenOpt_Default
    g_DefaultCC = LLVMX86StdcallCallConv
End Sub

Public Function CommandParser(ByVal Argc As Long, Argv() As String) As Long
    Dim i As Long
    For i = 1 To Argc - 1
        Select Case Argv(i)
        Case "-?"
            g_h = True
            Exit Function
        Case "-o"
            i = i + 1
            If i > Argc Then
                GeneralError "missing output file"
                GoTo ErrLine
            End If
            ReDim Preserve g_OutputFile(1 To g_FileCount) As String
            g_OutputFile(g_FileCount) = Argv(i)
        Case "-Gd"
            g_DefaultCC = LLVMCCallConv
        Case "-Gr"
            g_DefaultCC = LLVMX86FastcallCallConv
        Case "-Gz"
            g_DefaultCC = LLVMX86StdcallCallConv
        Case "-Gr"
        Case "-Gz"
        Case "-Vd"
            g_Vd = True
        Case "-BC"
            g_OutputType = OT_BITCODE
        Case "-LL"
            g_OutputType = OT_LLVMFILE
        Case "-AS"
            g_OutputType = OT_ASSEMBLY
        Case "-OB"
            g_OutputType = OT_OBJECT
        Case "-O0"
            g_OptLevel = LLVMCodeGenOpt_None
        Case "-O1"
            g_OptLevel = LLVMCodeGenOpt_Less
        Case "-O2"
            g_OptLevel = LLVMCodeGenOpt_Default
        Case "-O3"
            g_OptLevel = LLVMCodeGenOpt_Aggressive
        Case Else
            If Left$(Argv(i), 1) = "-" Then
                GeneralError "illegal option " & Argv(i)
                GoTo ErrLine
            Else
                g_FileCount = g_FileCount + 1
                ReDim Preserve g_FileList(1 To g_FileCount) As String
                g_FileList(g_FileCount) = Argv(i)
            End If
        End Select
    Next
    If g_FileCount <> 0 Then ReDim Preserve g_OutputFile(1 To g_FileCount) As String 'same bound to g_FileList
    Exit Function
ErrLine:
    CommandParser = -1
End Function

Public Sub ConsoleInit()
#If ENABLE_IDE_DEBUG Then
    If g_InIDE Then
        hStdOut = CreateFile("output.tmp", FILE_ALL_ACCESS, 0, ByVal 0, CREATE_ALWAYS, FILE_FLAG_DELETE_ON_CLOSE, 0)
        SetStdHandle STD_OUTPUT_HANDLE, hStdOut
        SetStdHandle STD_ERROR_HANDLE, hStdOut
    Else
#End If
        hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
        AllocConsole
#If ENABLE_IDE_DEBUG Then
    End If
#End If
    PrintLine "Extended BASIC Compiler v" & CStr(App.Major) & "." & CStr(App.Minor) & " r" & App.Revision & " (" & VERSION & ") license under GNU GPLv3"
End Sub

Public Sub ConsoleExit()
#If ENABLE_IDE_DEBUG Then
    If g_InIDE Then
        CloseHandle hStdOut
    Else
#End If
        FreeConsole
#If ENABLE_IDE_DEBUG Then
    End If
#End If
End Sub

Public Function PrintLine(s As String) As Long
    PrintLine = WriteFile2(hStdOut, s & vbCrLf, Len(s) + 2, 0, ByVal 0)
#If ENABLE_IDE_DEBUG Then
    PrintOutputFile
#End If
End Function

#If ENABLE_IDE_DEBUG Then
Public Sub PrintOutputFile()
    Static lp As Currency
    Dim lp2 As Currency
    Dim m As Long
    Dim s As String
    If g_InIDE Then
        GetFileSizeEx2 hStdOut, lp2
        If lp2 > lp Then
            m = (lp2 - lp) * 10000@
            s = Space(m \ 2 + 1)
            SetFilePointerEx hStdOut, lp, ByVal 0, SEEK_SET
            ReadFile hStdOut, ByVal StrPtr(s), m, 0, ByVal 0
            s = StrConv(s, vbUnicode)
            Debug.Print s
            lp = lp2
        End If
    End If
End Sub
#End If

Public Sub PrintHelp(ByVal s1 As String, ByVal s2 As String)
    PrintLine "  " & Format$(s1, "!@@@@@@@@@@@@@@@@@@@@@@@@") & "- " & s2
End Sub

Public Sub ShowHelp()
    PrintLine "Usage: EB [options] <files>"
    PrintLine ""
    PrintLine "Options:"
    PrintLine ""
    PrintLine "General"
    PrintHelp "-?", "Show this help (use this option exclusively)"
    PrintLine ""
    PrintLine "Compilation"
    PrintHelp "-Gd", "Use 'cdecl' as default calling convention"
    PrintHelp "-Gr", "Use 'x86_fastcall' as default calling convention"
    PrintHelp "-Gz", "Use 'x86_stdcall' as default calling convention (default)"
    PrintLine ""
    PrintLine "Output"
    PrintHelp "-o", "Name output file (follow by each input file, optional)"
    PrintHelp "-BC", "Output BitCode file"
    PrintHelp "-LL", "Output LLVM file"
    PrintHelp "-AS", "Output Assembly file"
    PrintHelp "-OB", "Output Object file"
    PrintLine ""
    PrintLine "Optimization"
    PrintHelp "-O0", "None optimization"
    PrintHelp "-O1", "Less optimization"
    PrintHelp "-O2", "Default optimization (default)"
    PrintHelp "-O3", "Aggressive optimization"
    PrintLine ""
    PrintLine "LLVM Options"
    PrintHelp "-Vd", "Disable LLVM Verifications (May cause error)"
End Sub

Public Sub PrintCounts()
    PrintLine vbNullString
    PrintLine CStr(WarningCount) & " warning(s), " & CStr(ErrorCount) & " error(s)"
End Sub

Public Sub PrintWarning(fname As String, ByVal Line As Long, s As String)
    PrintLine fname & "(" & CStr(Line) & ") WARNING: " & s
    WarningCount = WarningCount + 1
End Sub

Public Sub PrintError(fname As String, ByVal Line As Long, s As String)
    PrintLine fname & "(" & CStr(Line) & ") ERROR: " & s
    ErrorCount = ErrorCount + 1
End Sub

Public Sub GeneralError(s As String)
    PrintLine "Error: " & s
End Sub

Public Sub FatalError(s As String)
    PrintLine "Fatal Error: " & s
    End
End Sub

