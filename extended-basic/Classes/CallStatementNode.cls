VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallStatementNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

'on create
Dim m_Loc As Location

Dim m_Callee As String
Dim m_ArgList As ArgListNode

'on codegen
Dim m_Func As FunctionNode
Dim m_hValue As Long

Friend Function Create(Loc As Location, Callee As String, ByVal ArgList As ArgListNode) As Long
    Dim i As Long
    m_Loc = Loc
    m_Callee = Callee
    Set m_ArgList = ArgList
End Function

Friend Sub Codegen(ByVal C As Context)
    Dim hArgs() As Long
    If m_Func Is Nothing Then Set m_Func = C.CurrentFile.SymTable(m_Callee)
    If m_Func Is Nothing Then
        C.RaiseError "undefined function '" & m_Callee & "'"
        If C.UnContinuableError Then Exit Sub
    End If
    m_ArgList.Codegen C, m_Func.Proto
    If m_ArgList.Count > 0 Then
        hArgs = m_ArgList.hValues
        m_hValue = LLVMBuildCall(C.hBuilder, m_Func.hValue, hArgs(1), m_ArgList.Count, StrPtrA("calltmp"))
    Else
        m_hValue = LLVMBuildCall(C.hBuilder, m_Func.hValue, 0, 0, StrPtrA("calltmp"))
    End If
    LLVMSetInstructionCallConv m_hValue, LLVMGetFunctionCallConv(m_Func.hValue)
End Sub

Private Function IASTNode_Codegen(ByVal C As Context) As Long
    If C.Step And cg_all Then
        If m_hValue = 0 Then
            Codegen C
        End If
        IASTNode_Codegen = m_hValue
    End If
End Function

Private Function IASTNode_GetDescType(ByVal C As Context) As TypeNode
    Dim s As String
    If m_Func Is Nothing Then Set m_Func = C.CurrentFile.SymTable(m_Callee)
    If m_Func Is Nothing Then
        C.RaiseError "undefined function '" & m_Callee & "'"
        If C.UnContinuableError Then Exit Function
    End If
    If m_Func.Attr And mt_function Then
        Set IASTNode_GetDescType = m_Func.RetVal.CodegenGetType(C)
    End If
End Function

Private Function IASTNode_GetNodeType() As NODETYPE
    IASTNode_GetNodeType = nt_callstatement
End Function

Private Function IASTNode_IsConstant(ByVal C As Context) As Boolean
    IASTNode_IsConstant = False
End Function

Private Function IASTNode_Reverse(ByVal TabNum As Long) As String
    Dim Node As IASTNode
    Set Node = m_ArgList
    IASTNode_Reverse = m_Callee & Node.Reverse(TabNum)
End Function
