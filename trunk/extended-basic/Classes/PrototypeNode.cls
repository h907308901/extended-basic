VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrototypeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

'on create
Dim m_Loc As Location

Dim m_ArgName() As DimNode, m_ArgAttr() As ARGUMENTTYPE, m_ArgCount As Long '1-based
'on codegen
Dim m_hValue As Long

Friend Function Create(Loc As Location) As Long
    m_Loc = Loc
End Function

Friend Sub AddParam(ByVal ArgName As DimNode, ByVal ArgAttr As ARGUMENTTYPE)
    m_ArgCount = m_ArgCount + 1
    ReDim Preserve m_ArgName(1 To m_ArgCount) As DimNode
    ReDim Preserve m_ArgAttr(1 To m_ArgCount) As ARGUMENTTYPE
    Set m_ArgName(m_ArgCount) = ArgName
    m_ArgAttr(m_ArgCount) = ArgAttr
End Sub

Friend Property Get ArgCount() As Long
    ArgCount = m_ArgCount
End Property

Friend Property Get ArgName(ByVal Index As Long) As DimNode
    Set ArgName = m_ArgName(Index)
End Property

Friend Property Get ArgAttr(ByVal Index As Long) As ARGUMENTTYPE
    ArgAttr = m_ArgAttr(Index)
End Property

Friend Function Codegen(ByVal C As Context, ByVal hRetType As Long) As Long
    Dim hType() As Long, i As Long
    If m_hValue = 0 Then
        If m_ArgCount = 0 Then
            m_hValue = LLVMFunctionType(hRetType, 0, 0, 0)
        Else
            ReDim hType(1 To m_ArgCount + 2) As Long
            For i = 1 To m_ArgCount
                m_ArgName(i).CodegenArg C
                If C.CurrentFunction.SymTable.Exists(m_ArgName(i).Name) Then
                    C.RaiseError "'" & m_ArgName(i) & "' already exists"
                    If C.UnContinuableError Then Exit Function
                End If
                C.CurrentFunction.SymTable.Add m_ArgName(i).Name, m_ArgName(i)
                hType(i) = m_ArgName(i).CodegenGetType(C).Handle
            Next
            m_hValue = LLVMFunctionType(hRetType, hType(1), m_ArgCount, 0)
        End If
    End If
    Codegen = m_hValue
End Function

Private Function IASTNode_Codegen(ByVal C As Context) As Long
    ''' TODO:
End Function

Private Function IASTNode_GetDescType(ByVal C As Context) As TypeNode

End Function

Private Function IASTNode_GetNodeType() As NODETYPE
    IASTNode_GetNodeType = nt_prototype
End Function

Private Function IASTNode_IsConstant(ByVal C As Context) As Boolean

End Function

Private Function IASTNode_Reverse(ByVal TabNum As Long) As String
    Dim s As String
    Dim i As Long
    Dim Node As IASTNode
    s = "("
    For i = 1 To m_ArgCount
        If i > 1 Then
            s = s & ", "
        End If
        If m_ArgAttr(i) And at_byval Then
            s = s & "ByVal "
        Else
            s = s & "ByRef "
        End If
        s = s & m_ArgName(i).Name & " As " & m_ArgName(i).TypeName
    Next
    s = s & ")"
    IASTNode_Reverse = s
End Function
