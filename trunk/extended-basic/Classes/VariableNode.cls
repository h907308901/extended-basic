VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VariableNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

'on create
Dim m_Loc As Location
Dim m_VarName As String

'on codegen
Dim m_VarRef As DimNode
Dim m_hValue As Long

Friend Function Create(Loc As Location, Token As TOKEN_TYPE) As Long
    m_Loc = Loc
   m_VarName = Token.Identifier
End Function

Friend Property Get VarRef() As DimNode
    Set VarRef = m_VarRef
End Property

Friend Property Get TypeRef() As TypeNode
    Dim Var As DimNode
    Set Var = m_VarRef
    Set TypeRef = TypeByName(Var.TypeName)
    ''' TODO: if undefined
End Property

Friend Property Get Name() As String
    Name = m_VarName
End Property

Friend Sub Codegen(ByVal C As Context)
    Dim Node As IASTNode
    SeekVarRef C
    Set Node = m_VarRef
    If m_VarRef.ArgIndex >= 0 Then
        m_hValue = LLVMGetParam(C.CurrentFunction.hValue, m_VarRef.ArgIndex)
    Else
        m_hValue = LLVMBuildLoad(C.hBuilder, Node.Codegen(C), StrPtrA(VarRef.Name))
    End If
End Sub

Private Sub SeekVarRef(ByVal C As Context)
    Dim Func As FunctionNode, File As SourceFile
    Dim ST As Dictionary
    Dim Node As IASTNode
    If m_VarRef Is Nothing Then
        Set Func = C.CurrentFunction
        If Not Func Is Nothing Then Set ST = Func.SymTable
        If Not ST Is Nothing Then Set Node = ST(m_VarName)
        If Node Is Nothing Then
            Set File = C.CurrentFile
            If Not File Is Nothing Then Set ST = File.SymTable
            If Not ST Is Nothing Then Set Node = ST(m_VarName)
            If Node Is Nothing Then
                C.RaiseError "undefined variable '" & m_VarName & "'"
                If C.UnContinuableError Then Exit Sub
            End If
        End If
        Set m_VarRef = Node
    End If
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
    SeekVarRef C
    Set IASTNode_GetDescType = m_VarRef.CodegenGetType(C)
End Function

Private Function IASTNode_GetNodeType() As NODETYPE
    IASTNode_GetNodeType = nt_variable
End Function

Private Function IASTNode_IsConstant(ByVal C As Context) As Boolean
    SeekVarRef C
    IASTNode_IsConstant = m_VarRef.IsConstant(C)
End Function

Private Function IASTNode_Reverse(ByVal TabNum As Long) As String
    IASTNode_Reverse = m_VarName
End Function
