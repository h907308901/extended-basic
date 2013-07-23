VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgListNode"
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

'on codegen
Dim m_hValues() As Long '1-based

Friend Function Create(Loc As Location) As Long
    m_Loc = Loc
End Function

Friend Sub AddSubNode(ByVal SubNode As IASTNode)
    m_SubNodeCount = m_SubNodeCount + 1
    ReDim Preserve m_SubNode(1 To m_SubNodeCount) As IASTNode
    Set m_SubNode(m_SubNodeCount) = SubNode
End Sub

Friend Property Get Count() As Long
    Count = m_SubNodeCount
End Property

Friend Function hValues() As Long()
    hValues = m_hValues
End Function

Friend Sub Codegen(ByVal C As Context, ByVal Proto As PrototypeNode)
    Dim i As Long
    If m_SubNodeCount <> Proto.ArgCount Then
        C.RaiseError "parameter count does not match target function", , True
        If C.UnContinuableError Then Exit Sub
    End If
    If m_SubNodeCount <> 0 Then
        ReDim m_hValues(1 To m_SubNodeCount) As Long
        For i = 1 To m_SubNodeCount
            Proto.ArgName(i).CodegenArg C
            m_hValues(i) = m_SubNode(i).GetDescType(C).BuildTypeConversion(m_SubNode(i).Codegen(C), C, Proto.ArgName(i).Ty, m_SubNode(i).IsConstant(C))
        Next
    End If
End Sub

Private Function IASTNode_Codegen(ByVal C As Context) As Long
    
End Function

Private Function IASTNode_GetDescType(ByVal C As Context) As TypeNode

End Function

Private Function IASTNode_GetNodeType() As NODETYPE
    IASTNode_GetNodeType = nt_arglist
End Function

Private Function IASTNode_IsConstant(ByVal C As Context) As Boolean

End Function

Private Function IASTNode_Reverse(ByVal TabNum As Long) As String
    Dim s As String
    Dim i As Long
    For i = 1 To m_SubNodeCount
        If i = 1 Then
            s = "("
        Else
            s = s & ", "
        End If
        s = s & m_SubNode(i).Reverse(TabNum)
    Next
    s = s & ")"
    IASTNode_Reverse = s
End Function
