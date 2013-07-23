VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Context"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Step As CGSTEP
Dim m_hBuilder As Long
Dim m_hModule As Long
Dim m_Loc As Location
Dim m_UnContinuableError As Boolean
Dim m_File As SourceFile
Dim m_Function As FunctionNode

Friend Property Get Step() As CGSTEP
    Step = m_Step
End Property

Friend Property Let Step(ByVal Value As CGSTEP)
    m_Step = Value
End Property

Friend Sub RaiseError(ByVal s As String, Optional Warning As Boolean, Optional BreakError As Boolean)
    If Warning Then
        PrintWarning m_Loc.File, m_Loc.Line, s
    Else
        PrintError m_Loc.File, m_Loc.Line, s
        m_UnContinuableError = BreakError
    End If
End Sub

Friend Property Get UnContinuableError() As Boolean
    UnContinuableError = m_UnContinuableError
End Property

Friend Sub EnterFile(ByVal F As SourceFile)
    Dim s As String
    Set m_File = F
    m_hModule = F.hModule
End Sub

Friend Sub ExitFile()
    Set m_File = Nothing
    m_hModule = 0
End Sub

Friend Sub EnterFunction(ByVal F As FunctionNode)
    Set m_Function = F
End Sub

Friend Sub ExitFunction()
    Set m_Function = Nothing
End Sub

Friend Property Get hBuilder() As Long
    hBuilder = m_hBuilder
End Property

Friend Property Get hModule() As Long
    hModule = m_hModule
End Property

Friend Property Let Loc(l As Location)
    m_Loc = l
End Property

Friend Property Get Loc() As Location
    Loc = m_Loc
End Property

Friend Property Get CurrentFile() As SourceFile
    Set CurrentFile = m_File
End Property

Friend Property Get CurrentFunction() As FunctionNode
    Set CurrentFunction = m_Function
End Property

Private Sub Class_Initialize()
    m_hBuilder = LLVMCreateBuilder()
End Sub

Private Sub Class_Terminate()
    LLVMDisposeBuilder m_hBuilder
End Sub
