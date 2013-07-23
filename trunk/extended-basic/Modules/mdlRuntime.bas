Attribute VB_Name = "mdlRuntime"
Option Explicit

'****************************************
' Runtime
' operator, standard method support
'****************************************

Dim m_Op As Dictionary
Public UnContinuableError As Boolean

Public Sub RuntimeInit()
    Set m_Op = New Dictionary
    ''' Note: call LexerInit before calling RuntimeInit
    ' ----------
    AddPredOp "=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "+=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "-=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "*=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "/=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "^=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "&=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "|=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "^|=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "<<=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp ">>=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "<<<=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp ">>>=", of_binop Or of_make, 0
    ''' TODO:
    ' ----------
    AddPredOp "~", of_unaryl, 0
    'unary usages
    AddPredUsage "~", vt_ulonglong, 0, vt_ulonglong, sm_not
    AddPredUsage "~", vt_uint, 0, vt_uint, sm_not
    AddPredUsage "~", vt_ushort, 0, vt_ushort, sm_not
    AddPredUsage "~", vt_ubyte, 0, vt_ubyte, sm_not
    AddPredUsage "~", vt_longlong, 0, vt_longlong, sm_not
    AddPredUsage "~", vt_int, 0, vt_int, sm_not
    AddPredUsage "~", vt_short, 0, vt_short, sm_not
    AddPredUsage "~", vt_byte, 0, vt_byte, sm_not
    AddPredUsage "~", vt_boolean, 0, vt_boolean, sm_not
    ' ----------
    AddPredOp "==", of_binop, 5
    ''' TODO:
    ' ----------
    AddPredOp "<", of_binop, 5
    ''' TODO:
    ' ----------
    AddPredOp ">", of_binop, 5
    ''' TODO:
    ' ----------
    AddPredOp "<=", of_binop, 5
    ''' TODO:
    ' ----------
    AddPredOp ">=", of_binop, 5
    ''' TODO:
    ' ----------
    AddPredOp "<>", of_binop, 5
    ''' TODO:
    ' ----------
    AddPredOp "&&", of_binop, 1
    ''' TODO:
    ' ----------
    AddPredOp "||", of_binop, 1
    ''' TODO:
    ' ----------
    AddPredOp "+", of_binop Or of_unaryl, 10
    'binary usages
    AddPredUsage "+", vt_double, vt_double, vt_double, sm_fadd
    AddPredUsage "+", vt_float, vt_float, vt_float, sm_fadd
    AddPredUsage "+", vt_ulonglong, vt_ulonglong, vt_ulonglong, sm_add
    AddPredUsage "+", vt_uint, vt_uint, vt_uint, sm_add
    AddPredUsage "+", vt_ushort, vt_ushort, vt_ushort, sm_add
    AddPredUsage "+", vt_ubyte, vt_ubyte, vt_ubyte, sm_add
    AddPredUsage "+", vt_longlong, vt_longlong, vt_longlong, sm_add
    AddPredUsage "+", vt_int, vt_int, vt_int, sm_add
    AddPredUsage "+", vt_short, vt_short, vt_short, sm_add
    AddPredUsage "+", vt_byte, vt_byte, vt_byte, sm_add
    'unary usages
    AddPredUsage "+", vt_double, 0, vt_double, sm_retval
    AddPredUsage "+", vt_float, 0, vt_float, sm_retval
    AddPredUsage "+", vt_ulonglong, 0, vt_ulonglong, sm_retval
    AddPredUsage "+", vt_uint, 0, vt_uint, sm_retval
    AddPredUsage "+", vt_ushort, 0, vt_ushort, sm_retval
    AddPredUsage "+", vt_ubyte, 0, vt_ubyte, sm_retval
    AddPredUsage "+", vt_longlong, 0, vt_longlong, sm_retval
    AddPredUsage "+", vt_int, 0, vt_int, sm_retval
    AddPredUsage "+", vt_short, 0, vt_short, sm_retval
    AddPredUsage "+", vt_byte, 0, vt_byte, sm_retval
    ''' TODO: string
    ' ----------
    AddPredOp "-", of_binop Or of_unaryl, 10
    'binary usages
    AddPredUsage "-", vt_double, vt_double, vt_double, sm_fsub
    AddPredUsage "-", vt_float, vt_float, vt_float, sm_fsub
    AddPredUsage "-", vt_ulonglong, vt_ulonglong, vt_ulonglong, sm_sub
    AddPredUsage "-", vt_uint, vt_uint, vt_uint, sm_sub
    AddPredUsage "-", vt_ushort, vt_ushort, vt_ushort, sm_sub
    AddPredUsage "-", vt_ubyte, vt_ubyte, vt_ubyte, sm_sub
    AddPredUsage "-", vt_longlong, vt_longlong, vt_longlong, sm_sub
    AddPredUsage "-", vt_int, vt_int, vt_int, sm_sub
    AddPredUsage "-", vt_short, vt_short, vt_short, sm_sub
    AddPredUsage "-", vt_byte, vt_byte, vt_byte, sm_sub
    'unary usages
    AddPredUsage "-", vt_double, 0, vt_double, sm_fneg
    AddPredUsage "-", vt_float, 0, vt_float, sm_fneg
    AddPredUsage "-", vt_ulonglong, 0, vt_ulonglong, sm_neg
    AddPredUsage "-", vt_uint, 0, vt_uint, sm_neg
    AddPredUsage "-", vt_ushort, 0, vt_ushort, sm_neg
    AddPredUsage "-", vt_ubyte, 0, vt_ubyte, sm_neg
    AddPredUsage "-", vt_longlong, 0, vt_longlong, sm_neg
    AddPredUsage "-", vt_int, 0, vt_int, sm_neg
    AddPredUsage "-", vt_short, 0, vt_short, sm_neg
    AddPredUsage "-", vt_byte, 0, vt_byte, sm_neg
    ' ----------
    AddPredOp "*", of_binop, 20
    'binary usages
    AddPredUsage "*", vt_double, vt_double, vt_double, sm_fmul
    AddPredUsage "*", vt_float, vt_float, vt_float, sm_fmul
    AddPredUsage "*", vt_ulonglong, vt_ulonglong, vt_ulonglong, sm_mul
    AddPredUsage "*", vt_uint, vt_uint, vt_uint, sm_mul
    AddPredUsage "*", vt_ushort, vt_ushort, vt_ushort, sm_mul
    AddPredUsage "*", vt_ubyte, vt_ubyte, vt_ubyte, sm_mul
    AddPredUsage "*", vt_longlong, vt_longlong, vt_longlong, sm_mul
    AddPredUsage "*", vt_int, vt_int, vt_int, sm_mul
    AddPredUsage "*", vt_short, vt_short, vt_short, sm_mul
    AddPredUsage "*", vt_byte, vt_byte, vt_byte, sm_mul
    ''' TODO: string
    ' ----------
    AddPredOp "/", of_binop, 20
    'binary usages
    AddPredUsage "/", vt_double, vt_double, vt_double, sm_fdiv
    AddPredUsage "/", vt_float, vt_float, vt_float, sm_fdiv
    AddPredUsage "/", vt_ulonglong, vt_ulonglong, vt_ulonglong, sm_udiv
    AddPredUsage "/", vt_uint, vt_uint, vt_uint, sm_udiv
    AddPredUsage "/", vt_ushort, vt_ushort, vt_ushort, sm_udiv
    AddPredUsage "/", vt_ubyte, vt_ubyte, vt_ubyte, sm_udiv
    AddPredUsage "/", vt_longlong, vt_longlong, vt_longlong, sm_div
    AddPredUsage "/", vt_int, vt_int, vt_int, sm_div
    AddPredUsage "/", vt_short, vt_short, vt_short, sm_div
    AddPredUsage "/", vt_byte, vt_byte, vt_byte, sm_div
    ' ----------
    AddPredOp "%", of_binop, 20
    'binary usages
    AddPredUsage "%", vt_ulonglong, vt_ulonglong, vt_ulonglong, sm_umod
    AddPredUsage "%", vt_uint, vt_uint, vt_uint, sm_umod
    AddPredUsage "%", vt_ushort, vt_ushort, vt_ushort, sm_umod
    AddPredUsage "%", vt_ubyte, vt_ubyte, vt_ubyte, sm_umod
    AddPredUsage "%", vt_longlong, vt_longlong, vt_longlong, sm_mod
    AddPredUsage "%", vt_int, vt_int, vt_int, sm_mod
    AddPredUsage "%", vt_short, vt_short, vt_short, sm_mod
    AddPredUsage "%", vt_byte, vt_byte, vt_byte, sm_mod
    ' ----------
    AddPredOp "^", of_binop, 30
    ''' TODO:
    ' ----------
    AddPredOp "&", of_binop, 40
    'binary usages
    AddPredUsage "&", vt_ulonglong, vt_ulonglong, vt_ulonglong, sm_and
    AddPredUsage "&", vt_uint, vt_uint, vt_uint, sm_and
    AddPredUsage "&", vt_ushort, vt_ushort, vt_ushort, sm_and
    AddPredUsage "&", vt_ubyte, vt_ubyte, vt_ubyte, sm_and
    AddPredUsage "&", vt_longlong, vt_longlong, vt_longlong, sm_and
    AddPredUsage "&", vt_int, vt_int, vt_int, sm_and
    AddPredUsage "&", vt_short, vt_short, vt_short, sm_and
    AddPredUsage "&", vt_byte, vt_byte, vt_byte, sm_and
    ''' TODO: string
    ' ----------
    AddPredOp "|", of_binop, 40
    'binary usages
    AddPredUsage "|", vt_ulonglong, vt_ulonglong, vt_ulonglong, sm_or
    AddPredUsage "|", vt_uint, vt_uint, vt_uint, sm_or
    AddPredUsage "|", vt_ushort, vt_ushort, vt_ushort, sm_or
    AddPredUsage "|", vt_ubyte, vt_ubyte, vt_ubyte, sm_or
    AddPredUsage "|", vt_longlong, vt_longlong, vt_longlong, sm_or
    AddPredUsage "|", vt_int, vt_int, vt_int, sm_or
    AddPredUsage "|", vt_short, vt_short, vt_short, sm_or
    AddPredUsage "|", vt_byte, vt_byte, vt_byte, sm_or
    ' ----------
    AddPredOp "^|", of_binop, 40
    'binary usages
    AddPredUsage "^|", vt_ulonglong, vt_ulonglong, vt_ulonglong, sm_xor
    AddPredUsage "^|", vt_uint, vt_uint, vt_uint, sm_xor
    AddPredUsage "^|", vt_ushort, vt_ushort, vt_ushort, sm_xor
    AddPredUsage "^|", vt_ubyte, vt_ubyte, vt_ubyte, sm_xor
    AddPredUsage "^|", vt_longlong, vt_longlong, vt_longlong, sm_xor
    AddPredUsage "^|", vt_int, vt_int, vt_int, sm_xor
    AddPredUsage "^|", vt_short, vt_short, vt_short, sm_xor
    AddPredUsage "^|", vt_byte, vt_byte, vt_byte, sm_xor
    ' ----------
    AddPredOp "<<", of_binop, 50
    'binary usages
    AddPredUsage "<<", vt_ulonglong, vt_ulonglong, vt_ulonglong, sm_shl
    AddPredUsage "<<", vt_uint, vt_uint, vt_uint, sm_shl
    AddPredUsage "<<", vt_ushort, vt_ushort, vt_ushort, sm_shl
    AddPredUsage "<<", vt_ubyte, vt_ubyte, vt_ubyte, sm_shl
    AddPredUsage "<<", vt_longlong, vt_longlong, vt_longlong, sm_shl
    AddPredUsage "<<", vt_int, vt_int, vt_int, sm_shl
    AddPredUsage "<<", vt_short, vt_short, vt_short, sm_shl
    AddPredUsage "<<", vt_byte, vt_byte, vt_byte, sm_shl
    ' ----------
    AddPredOp ">>", of_binop, 50
    'binary usages
    AddPredUsage ">>", vt_ulonglong, vt_ulonglong, vt_ulonglong, sm_ushr
    AddPredUsage ">>", vt_uint, vt_uint, vt_uint, sm_ushr
    AddPredUsage ">>", vt_ushort, vt_ushort, vt_ushort, sm_ushr
    AddPredUsage ">>", vt_ubyte, vt_ubyte, vt_ubyte, sm_ushr
    AddPredUsage ">>", vt_longlong, vt_longlong, vt_longlong, sm_shr
    AddPredUsage ">>", vt_int, vt_int, vt_int, sm_shr
    AddPredUsage ">>", vt_short, vt_short, vt_short, sm_shr
    AddPredUsage ">>", vt_byte, vt_byte, vt_byte, sm_shr
    ' ----------
    AddPredOp "<<<", of_binop, 50
    ''' TODO:
    ' ----------
    AddPredOp ">>>", of_binop, 50
    ''' TODO:
    ' ----------
    AddPredOp "++", of_unaryr, 0
    ''' TODO:
    ' ----------
    AddPredOp "--", of_unaryr, 0
    ''' TODO:
    ' ----------
End Sub

Public Sub RuntimeExit()
    Set m_Op = Nothing
End Sub

Private Sub AddPredOp(Name As String, ByVal Flags As OP_FLAGS, ByVal Prec As Long)
    Dim Op As New Operator
    Op.Create TokenBySymbol(Name), Flags, Prec
    AddOperator Op
End Sub

Private Sub AddPredUsage(Name As String, ByVal Operand1 As VarType, ByVal Operand2 As VarType, ByVal ReturnType As VarType, ByVal Method As STANDARD_METHOD)
    Dim Op As Operator
    Dim Usage As OP_USAGE
    If m_Op.Exists(Name) Then
        With Usage
            Set .OperandType1 = TypeByBasicType(Operand1)
            If Operand2 <> 0 Then
                Set .OperandType2 = TypeByBasicType(Operand2)
                .OperandCount = 2
            Else
                .OperandCount = 1
            End If
            Set .ResultType = TypeByBasicType(ReturnType)
            .Method = "!STANDARD!" & CStr(Method)
        End With
        Set Op = m_Op(Name)
        Op.AddUsage Usage
    End If
End Sub

Public Function OperatorByName(s As String) As Operator
    If m_Op.Exists(s) Then
        Set OperatorByName = m_Op(s)
    End If
End Function

Public Function AddOperator(ByVal Op As Operator) As Long
    Dim s As String
    s = GetTokenName(Op.Token)
    If m_Op.Exists(s) Then
        AddOperator = -1
        Exit Function
    End If
    m_Op.Add s, Op
End Function

Public Sub ErrorBreak()
    UnContinuableError = True
End Sub

Public Function InvokeStandardMethod(ByVal Method As STANDARD_METHOD, ByVal hBuilder As Long, ByVal Param1 As Long, ByVal Param2 As Long, ByVal Param3 As Long, ByVal Param4 As Long, ByVal IsConstant As Long) As Long
    Dim i As Long
    Select Case Method
    Case sm_retval
        i = Param1
    Case sm_add '(i,i)i
        If IsConstant Then
            i = LLVMConstAdd(Param1, Param2)
        Else
            i = LLVMBuildAdd(hBuilder, Param1, Param2, StrPtrA("addtmp"))
        End If
    Case sm_sub '(i,i)i
        If IsConstant Then
            i = LLVMConstSub(Param1, Param2)
        Else
            i = LLVMBuildSub(hBuilder, Param1, Param2, StrPtrA("subtmp"))
        End If
    Case sm_mul '(i,i)i
        If IsConstant Then
            i = LLVMConstMul(Param1, Param2)
        Else
            i = LLVMBuildMul(hBuilder, Param1, Param2, StrPtrA("multmp"))
        End If
    Case sm_div '(i,i)i
        If IsConstant Then
            i = LLVMConstSDiv(Param1, Param2)
        Else
            i = LLVMBuildSDiv(hBuilder, Param1, Param2, StrPtrA("divtmp"))
        End If
    Case sm_mod '(i,i)i
        If IsConstant Then
            i = LLVMConstSRem(Param1, Param2)
        Else
            i = LLVMBuildSRem(hBuilder, Param1, Param2, StrPtrA("modtmp"))
        End If
    Case sm_fadd '(f,f)f
        If IsConstant Then
            i = LLVMConstFAdd(Param1, Param2)
        Else
            i = LLVMBuildFAdd(hBuilder, Param1, Param2, StrPtrA("faddtmp"))
        End If
    Case sm_fsub '(f,f)f
        If IsConstant Then
            i = LLVMConstFSub(Param1, Param2)
        Else
            i = LLVMBuildFSub(hBuilder, Param1, Param2, StrPtrA("fsubtmp"))
        End If
    Case sm_fmul '(f,f)f
        If IsConstant Then
            i = LLVMConstFMul(Param1, Param2)
        Else
            i = LLVMBuildFMul(hBuilder, Param1, Param2, StrPtrA("fmultmp"))
        End If
    Case sm_fdiv '(f,f)f
        If IsConstant Then
            i = LLVMConstFDiv(Param1, Param2)
        Else
            i = LLVMBuildFDiv(hBuilder, Param1, Param2, StrPtrA("fdivtmp"))
        End If
    Case sm_udiv '(u,u)u
        If IsConstant Then
            i = LLVMConstUDiv(Param1, Param2)
        Else
            i = LLVMBuildUDiv(hBuilder, Param1, Param2, StrPtrA("udivtmp"))
        End If
    Case sm_umod '(u,u)u
        If IsConstant Then
            i = LLVMConstURem(Param1, Param2)
        Else
            i = LLVMBuildURem(hBuilder, Param1, Param2, StrPtrA("umodtmp"))
        End If
    Case sm_neg '(i)i
        If IsConstant Then
            i = LLVMConstNeg(Param1)
        Else
            i = LLVMBuildNeg(hBuilder, Param1, StrPtrA("negtmp"))
        End If
    Case sm_fneg '(f)f
        If IsConstant Then
            i = LLVMConstFNeg(Param1)
        Else
            i = LLVMBuildFNeg(hBuilder, Param1, StrPtrA("fnegtmp"))
        End If
    Case sm_pow ''' TODO:
    Case sm_and '(i,i)i
        If IsConstant Then
            i = LLVMConstAnd(Param1, Param2)
        Else
            i = LLVMBuildAnd(hBuilder, Param1, Param2, StrPtrA("andtmp"))
        End If
    Case sm_or '(i,i)i
        If IsConstant Then
            i = LLVMConstOr(Param1, Param2)
        Else
            i = LLVMBuildOr(hBuilder, Param1, Param2, StrPtrA("ortmp"))
        End If
    Case sm_xor '(i,i)i
        If IsConstant Then
            i = LLVMConstXor(Param1, Param2)
        Else
            i = LLVMBuildXor(hBuilder, Param1, Param2, StrPtrA("xortmp"))
        End If
    Case sm_not '(i)i
        If IsConstant Then
            i = LLVMConstNot(Param1)
        Else
            i = LLVMBuildNot(hBuilder, Param1, StrPtrA("nottmp"))
        End If
    Case sm_shl '(i,i)i
        If IsConstant Then
            i = LLVMConstShl(Param1, Param2)
        Else
            i = LLVMBuildShl(hBuilder, Param1, Param2, StrPtrA("shltmp"))
        End If
    Case sm_shr '(i,i)i
        If IsConstant Then
            i = LLVMConstAShr(Param1, Param2)
        Else
            i = LLVMBuildAShr(hBuilder, Param1, Param2, StrPtrA("shrtmp"))
        End If
    Case sm_ushr '(u,u)u
        If IsConstant Then
            i = LLVMConstLShr(Param1, Param2)
        Else
            i = LLVMBuildLShr(hBuilder, Param1, Param2, StrPtrA("ushrtmp"))
        End If
    Case sm_rol ''' TODO:
    Case sm_ror ''' TODO:
    End Select
    InvokeStandardMethod = i
End Function
