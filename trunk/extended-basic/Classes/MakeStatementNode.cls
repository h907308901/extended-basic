VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MakeStatementNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

'on create
Dim m_Loc As Location

Dim m_LHS As IASTNode, m_RHS As IASTNode
'on codegen
Dim m_hLHS As Long
Dim m_hRHS As Long
Dim m_hValue As Long

Friend Function Create(Loc As Location, ByVal LHS As IASTNode, ByVal RHS As IASTNode) As Long
    m_Loc = Loc
    Set m_LHS = LHS
    Set m_RHS = RHS
End Function

Friend Sub Codegen(ByVal C As Context)
    Dim Var As VariableNode, Ty As TypeNode
    C.Loc = m_Loc
    If m_LHS.GetNodeType <> nt_variable Then
        C.RaiseError "lhs is not a variable"
        If C.UnContinuableError Then Exit Sub
    End If
    m_LHS.GetDescType C 'this step is only to make lhs get VarRef so that we can access it
    Set Var = m_LHS
    m_hLHS = Var.VarRef.VarHandle
    m_hRHS = m_RHS.GetDescType(C).BuildTypeConversion(m_RHS.Codegen(C), C, TypeByName(Var.VarRef.TypeName), m_RHS.IsConstant(C))
    m_hValue = LLVMBuildStore(C.hBuilder, m_hRHS, m_hLHS)
End Sub

Private Function IASTNode_Codegen(ByVal C As Context) As Long
    If C.Step And (cg_all Or cg_const) Then
        If m_hValue = 0 Then
            Codegen C
        End If
        IASTNode_Codegen = m_hValue
    End If
End Function

Private Function IASTNode_GetDescType(ByVal C As Context) As TypeNode
    Set IASTNode_GetDescType = m_LHS.GetDescType(C)
End Function

Private Function IASTNode_GetNodeType() As NODETYPE
    IASTNode_GetNodeType = nt_makestatement
End Function

Private Function IASTNode_IsConstant(ByVal C As Context) As Boolean

End Function

Private Function IASTNode_Reverse(ByVal TabNum As Long) As String
    IASTNode_Reverse = m_LHS.Reverse(TabNum) & " = " & m_RHS.Reverse(TabNum)
End Function
