VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FunctionNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

'on create
Dim m_Loc As Location

Dim m_Attr As METHODTYPE
Dim m_Name As String
Dim m_Proto As PrototypeNode 'can be nothing
Dim m_Statement As StatementListNode 'can be nothing
Dim m_ReturnVal As DimNode 'can be nothing
Dim m_SymTable As Dictionary
'on codegen
Dim m_hValue As Long
Dim m_hType As Long
Dim m_hRet As Long
Dim m_hRetType As Long
Dim m_hEntryBB As Long
Dim m_hEndBB As Long

Friend Function Create(Loc As Location, Name As String, ByVal Attributes As METHODTYPE, ByVal Prototype As PrototypeNode, ByVal ReturnVal As DimNode, ByVal StatementList As StatementListNode) As Long
    m_Loc = Loc
    m_Name = Name
    m_Attr = Attributes
    Set m_Proto = Prototype
    Set m_Statement = StatementList
    Set m_ReturnVal = ReturnVal
    m_SymTable.Add "Function", ReturnVal
End Function

Friend Property Get SymTable() As Dictionary
    Set SymTable = m_SymTable
End Property

Friend Property Get Attr() As METHODTYPE
    Attr = m_Attr
End Property

Friend Property Get Name() As String
    Name = m_Name
End Property

Friend Property Get Proto() As PrototypeNode
    Set Proto = m_Proto
End Property

Friend Property Get Statement() As StatementListNode
    Set Statement = m_Statement
End Property

Friend Property Get RetVal() As DimNode
    Set RetVal = m_ReturnVal
End Property

Friend Property Get hValue() As Long
    hValue = m_hValue
End Property

Friend Sub CodegenDef(ByVal C As Context)
    Dim i As Long, h As Long
    C.EnterFunction Me
    C.Loc = m_Loc
    If m_ReturnVal Is Nothing Then
        m_hRetType = LLVMVoidType
    Else
        m_hRetType = m_ReturnVal.CodegenGetType(C).Handle
    End If
    If m_Proto Is Nothing Then
        m_hType = LLVMFunctionType(m_hRetType, 0, 0, 0)
    Else
        m_hType = m_Proto.Codegen(C, m_hRetType)
    End If
    m_hValue = LLVMAddFunction(C.hModule, StrPtrA(m_Name), m_hType)
    If m_hValue = 0 Then
        C.RaiseError "create function '" & m_Name & "' failed", , True
        If C.UnContinuableError Then Exit Sub
    End If
    For i = 1 To m_Proto.ArgCount
        h = LLVMGetParam(m_hValue, i - 1)
        LLVMSetValueName h, StrPtrA(m_Proto.ArgName(i).Name)
    Next
    LLVMSetFunctionCallConv m_hValue, g_DefaultCC ''' TODO: custom callconv
    If m_Attr And mt_private Then
        LLVMSetLinkage m_hValue, LLVMPrivateLinkage
    Else
        LLVMSetLinkage m_hValue, LLVMExternalLinkage
    End If
    C.ExitFunction
End Sub

Private Sub Class_Initialize()
    Set m_SymTable = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set m_SymTable = Nothing
End Sub

Private Function IASTNode_Codegen(ByVal C As Context) As Long
    If C.Step And (cg_all Or cg_def) Then
        If m_hValue = 0 Then
            CodegenDef C
        End If
        IASTNode_Codegen = m_hValue
    End If
    If C.Step And cg_all Then
        If m_hEntryBB = 0 Then
            C.EnterFunction Me
            C.Loc = m_Loc
            m_hEntryBB = LLVMAppendBasicBlock(m_hValue, StrPtrA("entry"))
            If m_hEntryBB = 0 Then
                C.RaiseError "create basic block 'entry' failed in function'" & m_Name & "'", , True
                If C.UnContinuableError Then Exit Function
            End If
            LLVMPositionBuilderAtEnd C.hBuilder, m_hEntryBB
            If Not m_ReturnVal Is Nothing Then
                m_ReturnVal.Codegen C
                If C.UnContinuableError Then Exit Function
                m_hRet = m_ReturnVal.VarHandle
            End If
            m_Statement.Codegen C
            If C.UnContinuableError Then Exit Function
            m_hEndBB = LLVMAppendBasicBlock(m_hValue, StrPtrA("end"))
            If m_hEndBB = 0 Then
                C.RaiseError "create basic block 'end' failed in function'" & m_Name & "'", , True
                If C.UnContinuableError Then Exit Function
            End If
            LLVMPositionBuilderAtEnd C.hBuilder, m_hEntryBB
            LLVMBuildBr C.hBuilder, m_hEndBB
            LLVMPositionBuilderAtEnd C.hBuilder, m_hEndBB
            If m_hRet Then
                ''' TODO: return value type
                LLVMBuildRet C.hBuilder, LLVMBuildLoad(C.hBuilder, m_ReturnVal.VarHandle, StrPtrA("Ret"))
            Else
                LLVMBuildRetVoid C.hBuilder
            End If
            C.ExitFunction
        End If
    End If
End Function

Private Function IASTNode_GetDescType(ByVal C As Context) As TypeNode

End Function

Private Function IASTNode_GetNodeType() As NODETYPE
    IASTNode_GetNodeType = nt_function
End Function

Private Function IASTNode_IsConstant(ByVal C As Context) As Boolean

End Function

Private Function IASTNode_Reverse(ByVal TabNum As Long) As String
    Dim s As String
    Dim Node As IASTNode
    s = String$(TabNum, vbTab)
    If m_Attr And mt_public Then
        s = s & "Public "
    Else
        s = s & "Private "
    End If
    If m_Attr And mt_function Then
        s = s & "Function "
    Else
        s = s & "Sub"
    End If
    s = s & m_Name & " "
    If m_Proto Is Nothing Then
        s = s & "()"
    Else
        Set Node = m_Proto
        s = s & Node.Reverse(TabNum)
    End If
    If m_Attr And mt_function Then
        s = s & " As " & m_ReturnVal.TypeName
    End If
    s = s & vbCrLf
    If m_Statement Is Nothing Then
        s = s & String$(TabNum + 1, vbTab)
    Else
        Set Node = m_Statement
        s = s & Node.Reverse(TabNum + 1)
    End If
    If m_Attr And mt_function Then
        s = s & "End Function"
    Else
        s = s & "End Sub "
    End If
    s = s & vbCrLf
    IASTNode_Reverse = s
End Function
