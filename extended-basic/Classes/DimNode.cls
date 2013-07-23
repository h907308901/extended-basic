VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DimNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

'on create
Dim m_Loc As Location

Dim m_Name As String, m_Type As String
Dim m_ArgIndex As Long
Dim m_IsConstant As Boolean
'on codegen
Dim m_Ty As TypeNode
Dim m_hValue As Long
''' TODO: mutiple dim

Friend Function Create(Loc As Location, VarName As String, VarType As String, Optional ByVal ArgIndex As Long = -1, Optional ByVal IsConstant As Boolean) As Long
    m_Loc = Loc
    m_Name = VarName
    m_Type = VarType
    m_ArgIndex = ArgIndex
End Function

Friend Property Get Name() As String
    Name = m_Name
End Property

Friend Property Get TypeName() As String
    TypeName = m_Type
End Property

Friend Property Get Ty() As TypeNode
    Set Ty = m_Ty
End Property

Friend Property Get VarHandle() As Long
    VarHandle = m_hValue
End Property

Friend Property Get ArgIndex() As Long
    ArgIndex = m_ArgIndex
End Property

Friend Property Get IsConstant(ByVal C As Context) As Boolean
    IsConstant = m_IsConstant
End Property

Friend Function CodegenGetType(ByVal C As Context) As TypeNode
    If m_Ty Is Nothing Then
        Set m_Ty = TypeByName(m_Type)
        If m_Ty Is Nothing Then
            C.RaiseError "undefined type '" & m_Type & "'", , True
            Exit Function
        End If
    End If
    Set CodegenGetType = m_Ty
End Function

Friend Sub Codegen(ByVal C As Context)
    C.Loc = m_Loc
    CodegenGetType C
    ''' TODO: other types
    m_hValue = LLVMBuildAlloca(C.hBuilder, m_Ty.Handle, StrPtrA(m_Name))
End Sub

Friend Sub CodegenArg(ByVal C As Context)
    C.Loc = m_Loc
    Set m_Ty = TypeByName(m_Type)
    If m_Ty Is Nothing Then
        C.RaiseError "undefined type '" & m_Type & "'", , True
        Exit Sub
    End If
    ''' TODO: array and struct types, byref
End Sub

Private Function IASTNode_Codegen(ByVal C As Context) As Long
    If C.Step And (cg_all Or cg_def) Then
        If m_hValue = 0 Then
            Codegen C
        End If
        IASTNode_Codegen = m_hValue
    End If
End Function

Private Function IASTNode_GetDescType(ByVal C As Context) As TypeNode

End Function

Private Function IASTNode_GetNodeType() As NODETYPE
    IASTNode_GetNodeType = nt_dim
End Function

Private Function IASTNode_IsConstant(ByVal C As Context) As Boolean
    
End Function

Private Function IASTNode_Reverse(ByVal TabNum As Long) As String
    IASTNode_Reverse = "Dim " & m_Name & " As " & m_Type ''' TODO: mutiple dim
End Function
