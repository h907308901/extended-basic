VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Operator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Initialized As Boolean
Dim m_Token As TOKEN_ENUM 'must be defined
Dim m_Flags As OP_FLAGS
Dim m_Prec As Long ''' Note: precedence only avaliable for binop
Dim m_Usages() As OP_USAGE, m_UsageCount As Long '1-based
''' Note: traverse from 1 to ubound, take the usage if one operand meets it

Friend Function Create(ByVal Token As TOKEN_ENUM, ByVal Flags As OP_FLAGS, ByVal Precedence As Long) As Long
    If m_Initialized Then
        Create = -1
        Exit Function
    End If
    m_Token = Token
    m_Flags = Flags
    m_Prec = Precedence
    m_Initialized = True
End Function

Friend Sub AddUsage(Usage As OP_USAGE)
    m_UsageCount = m_UsageCount + 1
    ReDim Preserve m_Usages(1 To m_UsageCount) As OP_USAGE
    m_Usages(m_UsageCount) = Usage
End Sub

Friend Function Invoke(ByVal C As Context, ByVal OperandCount As Long, ByVal Operand1 As Long, ByVal Operand2 As Long, ByVal Type1 As TypeNode, ByVal Type2 As TypeNode, Optional ByVal IsConstant As Boolean) As Long
    Dim i As Long, b1 As Boolean, b2 As Boolean
    Dim DescTy1 As TypeNode, DescTy2 As TypeNode
    Dim h1 As Long, h2 As Long
    Dim a As Long
    For i = 1 To m_UsageCount
        If m_Usages(i).OperandCount = OperandCount Then
            Set DescTy1 = m_Usages(i).OperandType1
            If OperandCount > 1 Then Set DescTy2 = m_Usages(i).OperandType2
            b1 = Type1.IsEqualTo(DescTy1)
            If OperandCount > 1 Then b2 = Type2.IsEqualTo(DescTy2)
            If b1 Or b2 Then
                If b1 Then
                    h1 = Operand1
                Else
                    h1 = Type1.BuildTypeConversion(Operand1, C, DescTy1, IsConstant)
                End If
                If OperandCount > 1 Then
                    If b2 Then
                        h2 = Operand2
                    Else
                        h2 = Type2.BuildTypeConversion(Operand2, C, DescTy2, IsConstant)
                    End If
                End If
                If Left$(m_Usages(i).Method, 10) = "!STANDARD!" Then
                    a = Val(Right$(m_Usages(i).Method, Len(m_Usages(i).Method) - 10))
                    Invoke = InvokeStandardMethod(a, C.hBuilder, h1, h2, 0, 0, IsConstant)
                    Exit Function
                Else
                    ''' TODO:
                End If
            End If
        End If
    Next
End Function

Friend Function CalcDescType(ByVal OperandCount As Long, ByVal Type1 As TypeNode, ByVal Type2 As TypeNode) As TypeNode
    Dim i As Long, b1 As Boolean, b2 As Boolean
    Dim DescTy1 As TypeNode, DescTy2 As TypeNode
    For i = 1 To m_UsageCount
        If m_Usages(i).OperandCount = OperandCount Then
            Set DescTy1 = m_Usages(i).OperandType1
            If OperandCount > 1 Then Set DescTy2 = m_Usages(i).OperandType2
            b1 = Type1.IsEqualTo(DescTy1)
            If OperandCount > 1 Then b2 = Type2.IsEqualTo(DescTy2)
            If b1 Or b2 Then
                Set CalcDescType = m_Usages(i).ResultType
                Exit Function
            End If
        End If
    Next
End Function

Friend Property Get Token() As TOKEN_ENUM
    Token = m_Token
End Property

Friend Property Get Flags() As OP_FLAGS
    Flags = m_Flags
End Property

Friend Property Get Precedence() As Long
    Precedence = m_Prec
End Property

Friend Property Get UsageCount() As Long
    UsageCount = m_UsageCount
End Property

Friend Property Get Usage(ByVal Index As Long) As OP_USAGE
    If Index > 0 And Index <= m_UsageCount Then
        Usage = m_Usages(Index)
    End If
End Property
