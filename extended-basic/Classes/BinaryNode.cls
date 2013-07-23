VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

'on create
Dim m_Loc As Location

Dim m_LHS As IASTNode
Dim m_RHS As IASTNode
Dim m_Op As Long
'on codegen
Dim m_hLHS As Long
Dim m_hRHS As Long
Dim m_hValue As Long

Friend Function Create(Loc As Location, ByVal Op As Long, ByVal LHS As IASTNode, ByVal RHS As IASTNode) As Long
    m_Loc = Loc
    Set m_LHS = LHS
    Set m_RHS = RHS
    m_Op = Op
End Function

Friend Sub Codegen(ByVal C As Context)
    Dim Op As Operator
    m_hLHS = m_LHS.Codegen(C)
    m_hRHS = m_RHS.Codegen(C)
    If m_hLHS = 0 Or m_hRHS = 0 Then Exit Sub
    Set Op = OperatorByName(GetTokenName(m_Op))
    If Not Op Is Nothing Then
        m_hValue = Op.Invoke(C, 2, m_hLHS, m_hRHS, m_LHS.GetDescType(C), m_RHS.GetDescType(C), IASTNode_IsConstant(C))
    End If
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
    Dim Op As Operator
    Set Op = OperatorByName(GetTokenName(m_Op))
    If Not Op Is Nothing Then
        Set IASTNode_GetDescType = Op.CalcDescType(2, m_LHS.GetDescType(C), m_RHS.GetDescType(C))
    End If
End Function

Private Function IASTNode_GetNodeType() As NODETYPE
    IASTNode_GetNodeType = nt_binary
End Function

Private Function IASTNode_IsConstant(ByVal C As Context) As Boolean
    IASTNode_IsConstant = m_LHS.IsConstant(C) And m_RHS.IsConstant(C)
End Function

Private Function IASTNode_Reverse(ByVal TabNum As Long) As String
    IASTNode_Reverse = m_LHS.Reverse(TabNum) & " " & GetTokenName(m_Op) & " " & m_RHS.Reverse(TabNum)
End Function
