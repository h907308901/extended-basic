VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' Note: this class is not a sub class of IASTNode

'on create
Dim m_TypeName As String
Dim m_Initialized As Boolean
Dim m_Size As Long
'for basic types
Dim m_BasicType As VarType
Dim m_BasicTypeFlags As TypeFlags
'for struct
''' TODO:
'for enum
''' TODO:
Dim m_hType As Long

Friend Function CreateBasicType(TypeName As String, ByVal BType As VarType) As Long
    ''' Note: this function is called by DefineBasicType (mdlTypeMgr), do not call by yourself, multiple same basic type may cause problem
    If m_Initialized Then
        CreateBasicType = -1
        Exit Function
    End If
    m_TypeName = TypeName
    m_BasicType = BType
    m_Initialized = True
    Select Case BType
    Case vt_byte, vt_ubyte
        m_hType = LLVMInt8Type
    Case vt_short, vt_ushort
        m_hType = LLVMInt16Type
    Case vt_int, vt_uint
        m_hType = LLVMInt32Type
    Case vt_longlong, vt_ulonglong
        m_hType = LLVMInt64Type
    Case vt_float
        m_hType = LLVMFloatType
    Case vt_double
        m_hType = LLVMDoubleType
    End Select
    InitializeTypeProperties
End Function

Friend Function BuildTypeConversion(ByVal hValue As Long, ByVal C As Context, ByVal DescType As TypeNode, Optional ByVal IsConstant As Boolean) As Long
    Dim DescBasicType As VarType
    If Not m_Initialized Then
        BuildTypeConversion = -1
        Exit Function
    End If
    If IsEqualTo(DescType) Then
        BuildTypeConversion = hValue
        Exit Function
    End If
    DescBasicType = DescType.BasicType
    If DescBasicType <> vt_struct And DescBasicType <> vt_string Then
        BuildTypeConversion = BasicTypeConversion(hValue, C.hBuilder, Me, DescType, IsConstant)
    Else
        ''' TODO:
    End If
End Function

Friend Function IsEqualTo(ByVal Ty As TypeNode) As Boolean
    IsEqualTo = (m_TypeName = Ty.Name)
End Function

Friend Property Get Name() As String
    Name = m_TypeName
End Property

Friend Property Get Size() As Long
    Size = m_Size
End Property

Friend Property Get Flags() As TypeFlags
    Flags = m_BasicTypeFlags
End Property

Friend Property Get BasicType() As VarType
    BasicType = m_BasicType
End Property

Friend Property Get Handle() As Long
    Handle = m_hType
End Property

Private Sub InitializeTypeProperties()
    If m_BasicType <> vt_struct And m_BasicType <> vt_string Then
        m_Size = BasicTypeSize(m_BasicType)
        m_BasicTypeFlags = BasicTypeFlags(m_BasicType)
    Else
        ''' TODO:
    End If
End Sub
