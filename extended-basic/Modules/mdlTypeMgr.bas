Attribute VB_Name = "mdlTypeMgr"
Option Explicit

'****************************************
' TypeMgr
' type support
'****************************************

Public g_TypeTable As Dictionary
Dim BT_Type_Translation(1 To vt_max_value) As TypeNode 'basictype-type translation table

Public Sub TypeInit()
    Set g_TypeTable = New Dictionary
    DefineBasicType "Byte", vt_byte
    DefineBasicType "Long", vt_int
    DefineBasicType "Integer", vt_short
    DefineBasicType "LongLong", vt_longlong
    DefineBasicType "UnsignedByte", vt_ubyte
    DefineBasicType "UnsignedInteger", vt_ushort
    DefineBasicType "UnsignedLong", vt_uint
    DefineBasicType "UnsignedLongLong", vt_ulonglong
    DefineBasicType "Single", vt_float
    DefineBasicType "Double", vt_double
    DefineBasicType "Boolean", vt_boolean ''' TODO:
End Sub

Public Sub TypeExit()
    Set g_TypeTable = Nothing
End Sub

Public Function DefineBasicType(TypeName As String, ByVal BasicType As VarType) As TypeNode
    Dim Node As New TypeNode
    Node.CreateBasicType TypeName, BasicType
    g_TypeTable.Add TypeName, Node
    Set BT_Type_Translation(BasicType) = Node
    Set DefineBasicType = Node
End Function

Public Function TypeByName(s As String) As TypeNode
    Dim Node As TypeNode
    Set Node = g_TypeTable(s)
    Set TypeByName = Node
End Function

Public Function TypeByBasicType(ByVal BasicType As VarType) As TypeNode
    If BasicType > 0 And BasicType < vt_max_value Then
        Set TypeByBasicType = BT_Type_Translation(BasicType)
    End If
End Function

Public Function BasicTypeFlags(ByVal BasicType As VarType) As TypeFlags
    Select Case BasicType
    Case vt_byte, vt_short, vt_int, vt_longlong
        BasicTypeFlags = tf_signed
    Case vt_ubyte, vt_ushort, vt_uint, vt_ulonglong
        BasicTypeFlags = tf_unsigned
    Case vt_float, vt_double
        BasicTypeFlags = tf_float
    End Select
End Function

Public Function BasicTypeSize(ByVal BasicType As VarType) As Long
    Select Case BasicType
    Case vt_byte, vt_ubyte
        BasicTypeSize = 1
    Case vt_short, vt_ushort
        BasicTypeSize = 2
    Case vt_int, vt_uint, vt_float
        BasicTypeSize = 4
    Case vt_longlong, vt_ulonglong, vt_double
        BasicTypeSize = 8
    End Select
End Function

Public Function BasicTypeConversion(ByVal hValue As Long, ByVal hBuilder As Long, ByVal SrcType As TypeNode, ByVal DescType As TypeNode, ByVal IsConstant As Boolean) As Long
    ''' Note: this function is a internal function for TypeNode::BuildTypeConversion
    Dim namestr As Long
    namestr = StrPtrA("convtemp")
    Select Case DescType.Flags And &HF
    Case tf_signed, tf_unsigned
        Select Case SrcType.Flags
        Case tf_signed, tf_unsigned
            If DescType.Size <= SrcType.Size Then
                If IsConstant Then
                    BasicTypeConversion = LLVMConstIntCast(hValue, DescType.Handle, (DescType.Flags And &HF) = tf_unsigned)
                Else
                    BasicTypeConversion = LLVMBuildIntCast(hBuilder, hValue, DescType.Handle, namestr)
                End If
            ElseIf (SrcType.Flags And &HF) = tf_signed Then
                If IsConstant Then
                    BasicTypeConversion = LLVMConstSExt(hValue, DescType.Handle)
                Else
                    BasicTypeConversion = LLVMBuildSExt(hBuilder, hValue, DescType.Handle, namestr)
                End If
            Else
                If IsConstant Then
                    BasicTypeConversion = LLVMConstZExt(hValue, DescType.Handle)
                Else
                    BasicTypeConversion = LLVMBuildZExt(hBuilder, hValue, DescType.Handle, namestr)
                End If
            End If
        Case tf_float
            If (DescType.Flags And &HF) = tf_signed Then
                If IsConstant Then
                    BasicTypeConversion = LLVMConstFPToSI(hValue, DescType.Handle)
                Else
                    BasicTypeConversion = LLVMBuildFPToSI(hBuilder, hValue, DescType.Handle, namestr)
                End If
            Else
                If IsConstant Then
                    BasicTypeConversion = LLVMConstFPToUI(hValue, DescType.Handle)
                Else
                    BasicTypeConversion = LLVMBuildFPToUI(hBuilder, hValue, DescType.Handle, namestr)
                End If
            End If
        End Select
    Case tf_float
        Select Case SrcType.Flags And &HF
        Case tf_signed
            If IsConstant Then
                BasicTypeConversion = LLVMConstSIToFP(hValue, DescType.Handle)
            Else
                BasicTypeConversion = LLVMBuildSIToFP(hBuilder, hValue, DescType.Handle, namestr)
            End If
        Case tf_unsigned
            If IsConstant Then
                BasicTypeConversion = LLVMConstUIToFP(hValue, DescType.Handle)
            Else
                BasicTypeConversion = LLVMBuildUIToFP(hBuilder, hValue, DescType.Handle, namestr)
            End If
        Case tf_float
            If IsConstant Then
                BasicTypeConversion = LLVMConstFPCast(hValue, DescType.Handle)
            Else
                BasicTypeConversion = LLVMBuildFPCast(hBuilder, hValue, DescType.Handle, namestr)
            End If
        End Select
    End Select
End Function
