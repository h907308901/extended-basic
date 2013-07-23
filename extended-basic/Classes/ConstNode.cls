VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

'on create
Dim m_Loc As Location

Dim m_Token As TOKEN_TYPE
Dim m_Type As TypeNode
'on codegen
Dim m_hType As Long
Dim m_hValue As Long

Friend Function Create(Loc As Location, Token As TOKEN_TYPE, ByVal ConstType As TypeNode) As Long
    m_Loc = Loc
    m_Token = Token
    Set m_Type = ConstType
End Function

Friend Sub Codegen(ByVal C As Context)
    Select Case m_Type.BasicType
    Case vt_byte
        m_hType = LLVMInt8Type
        m_hValue = LLVMConstInt(m_hType, Val(m_Token.Identifier) * 0.0001@, 1)
    Case vt_short
        m_hType = LLVMInt16Type
        m_hValue = LLVMConstInt(m_hType, Val(m_Token.Identifier) * 0.0001@, 1)
    Case vt_int
        m_hType = LLVMInt32Type
        m_hValue = LLVMConstInt(m_hType, Val(m_Token.Identifier) * 0.0001@, 1)
    Case vt_longlong
        m_hType = LLVMInt64Type
        ''' TODO:
    Case vt_ubyte
        m_hType = LLVMInt8Type
        m_hValue = LLVMConstInt(m_hType, Val(m_Token.Identifier) * 0.0001@, 0)
    Case vt_ushort
        m_hType = LLVMInt16Type
        m_hValue = LLVMConstInt(m_hType, Val(m_Token.Identifier) * 0.0001@, 0)
    Case vt_uint
        m_hType = LLVMInt32Type
        m_hValue = LLVMConstInt(m_hType, Val(m_Token.Identifier) * 0.0001@, 0)
    Case vt_ulonglong
        m_hType = LLVMInt64Type
        ''' TODO:
    Case vt_string
        ''' TODO:
    Case vt_float
        m_hType = LLVMFloatType
        m_hValue = LLVMConstReal(m_hType, Val(m_Token.Identifier))
    Case vt_double
        m_hType = LLVMDoubleType
        m_hValue = LLVMConstReal(m_hType, Val(m_Token.Identifier))
    Case vt_struct
        ''' TODO:
    End Select
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
    Set IASTNode_GetDescType = m_Type
End Function

Private Function IASTNode_GetNodeType() As NODETYPE
    IASTNode_GetNodeType = nt_const
End Function

Private Function IASTNode_IsConstant(ByVal C As Context) As Boolean
    IASTNode_IsConstant = True
End Function

Private Function IASTNode_Reverse(ByVal TabNum As Long) As String
    IASTNode_Reverse = m_Token.Identifier
End Function
