VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatementListNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

'on create
Dim m_Loc As Location

Dim m_SubNode() As IASTNode '1-based
Dim m_SubNodeCount As Long

Friend Function Create(Loc As Location) As Long
    m_Loc = Loc
End Function

Friend Sub AddSubNode(ByVal SubNode As IASTNode)
    m_SubNodeCount = m_SubNodeCount + 1
    ReDim Preserve m_SubNode(1 To m_SubNodeCount) As IASTNode
    Set m_SubNode(m_SubNodeCount) = SubNode
End Sub

Public Sub Codegen(ByVal C As Context)
    Dim i As Long
    For i = 1 To m_SubNodeCount
        m_SubNode(i).Codegen C
    Next
End Sub

Private Function IASTNode_Codegen(ByVal C As Context) As Long
    If C.Step = cg_all Then
        Codegen C
    End If
End Function

Private Function IASTNode_GetDescType(ByVal C As Context) As TypeNode

End Function

Private Function IASTNode_GetNodeType() As NODETYPE
    IASTNode_GetNodeType = nt_statementlist
End Function

Private Function IASTNode_IsConstant(ByVal C As Context) As Boolean

End Function

Private Function IASTNode_Reverse(ByVal TabNum As Long) As String
    Dim s As String
    Dim i As Long
    For i = 1 To m_SubNodeCount
        s = s & String$(TabNum, vbTab)
        s = s & m_SubNode(i).Reverse(TabNum)
        s = s & vbCrLf
    Next
    IASTNode_Reverse = s
End Function
