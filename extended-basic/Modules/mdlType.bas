Attribute VB_Name = "mdlType"
Option Explicit

'****************************************
' Type
' all types and enums defined here
'****************************************

Public Enum VBCharType
    [EOF] = -1
    ["\0"] = 0
    ["\t"] = 9
    ["\n"] = 10
    ["\r"] = 13
    [" "] = 32
    ["!"] = 33
    ["""] = 34
    ["#"] = 35
    ["$"] = 36
    ["%"] = 37
    ["&"] = 38
    ["'"] = 39
    ["("] = 40
    [")"] = 41
    ["*"] = 42
    ["+"] = 43
    [","] = 44
    ["-"] = 45
    ["."] = 46
    ["/"] = 47
    ["0"] = 48
    ["1"] = 49
    ["2"] = 50
    ["3"] = 51
    ["4"] = 52
    ["5"] = 53
    ["6"] = 54
    ["7"] = 55
    ["8"] = 56
    ["9"] = 57
    ["col"] = 58                                            ':
    [";"] = 59
    ["<"] = 60
    ["="] = 61
    [">"] = 62
    ["?"] = 63
    ["@"] = 64
    ["a"] = 65
    ["b"] = 66
    ["c"] = 67
    ["d"] = 68
    ["e"] = 69
    ["f"] = 70
    ["g"] = 71
    ["h"] = 72
    ["i"] = 73
    ["j"] = 74
    ["k"] = 75
    ["l"] = 76
    ["m"] = 77
    ["n"] = 78
    ["o"] = 79
    ["p"] = 80
    ["q"] = 81
    ["r"] = 82
    ["s"] = 83
    ["t"] = 84
    ["u"] = 85
    ["v"] = 86
    ["w"] = 87
    ["x"] = 88
    ["y"] = 89
    ["z"] = 90
    ["lll"] = 91                                            '[
    ["\"] = 92
    ["rrr"] = 93                                            ']
    ["^"] = 94
    ["_"] = 95
    ["`"] = 96
    ["aa"] = 97
    ["bb"] = 98
    ["cc"] = 99
    ["dd"] = 100
    ["ee"] = 101
    ["ff"] = 102
    ["gg"] = 103
    ["hh"] = 104
    ["ii"] = 105
    ["jj"] = 106
    ["kk"] = 107
    ["ll"] = 108
    ["mm"] = 109
    ["nn"] = 110
    ["oo"] = 111
    ["pp"] = 112
    ["qq"] = 113
    ["rr"] = 114
    ["ss"] = 115
    ["tt"] = 116
    ["uu"] = 117
    ["vv"] = 118
    ["ww"] = 119
    ["xx"] = 120
    ["yy"] = 121
    ["zz"] = 122
    ["{"] = 123
    ["|"] = 124
    ["}"] = 125
    ["~"] = 126
End Enum

Public Enum TOKEN_ENUM
'basic types
    token_eof = -1
    token_crlf
    token_identifier
    token_integer
    token_float
    token_string
    token_unknown_symbol
    'reserved symbols
    token_reserved_symbol_start = 5000
    token_dot                                               '.
    token_comma                                             ',
    token_lbraket                                           '(
    token_rbraket                                           ')
'    'special symbols
'    token_dot                                               '.
'    token_comma                                             ',
'    token_lbraket                                           '(
'    token_rbraket                                           ')
'    token_make                                              '=
'    token_mplus                                             '+=
'    token_mminus                                            '-=
'    token_mmul                                              '*=
'    token_mdiv                                              '/=
'    token_mpower                                            '^=
'    token_mand                                              '&=
'    token_mor                                               '|=
'    token_mxor                                              '^|=
'    token_mshl                                              '<<=
'    token_mshr                                              '>>=
'    token_mrol                                              '<<<=
'    token_mror                                              '>>>=
'    token_not                                               '~ unary
'    'operation symbols
'    token_equal                                             '==
'    token_lt                                                '<
'    token_gt                                                '>
'    token_le                                                '<=
'    token_ge                                                '>=
'    token_ne                                                '<>
'    '----------
'    token_land                                              '&&
'    token_lor                                               '||
'    '----------
'    token_plus                                              '+, may be unary
'    token_minus                                             '-, may be unary
'    '----------
'    token_mul                                               '*
'    token_div                                               '/
'    token_mod                                               '%
'    '----------
'    token_power                                             '^
'    '----------
'    token_and                                               '&
'    token_or                                                '|
'    token_xor                                               '^|
'    '----------
'    token_shl                                               '<<
'    token_shr                                               '>>
'    token_rol                                               '<<<
'    token_ror                                               '>>>
'    '----------
'    token_plus1                                             '++
'    token_minus1                                            '--
    'keywords
    token_keyword_start = 10000
    keyword_as
    keyword_byref
    keyword_byval
    keyword_dim
    keyword_end
    keyword_function
    keyword_sub
    keyword_private
    keyword_public
    'symbols
    token_symbol_start = 20000
End Enum

Public Type TOKEN_TYPE
    TokenType As TOKEN_ENUM
    Line As Long
    Identifier As String
    IsError As Long
End Type

Public Type Location
    File As String
    Line As Long
End Type

Public Enum VarType
    vt_byte = 1
    vt_short
    vt_int
    vt_longlong
    vt_ubyte
    vt_ushort
    vt_uint
    vt_ulonglong
    vt_float
    vt_double
    vt_boolean ''' TODO:
    vt_string ''' TODO:
    vt_struct
    vt_max_value
End Enum

Public Enum TypeFlags
    'group 0x1~0xF
    tf_signed = 1
    tf_unsigned
    tf_float
End Enum

Public Enum OP_FLAGS
    of_unaryr = &H1 'right-handed unary, such as ++ --
    of_unaryl = &H2 'left-handed unary,such as + -
    of_binop = &H10 'binocular operator
    of_make = &H100 'make operator
End Enum

Public Type OP_USAGE
    OperandCount As Long
    OperandType1 As TypeNode
    OperandType2 As TypeNode 'optional
    ResultType As TypeNode
    Method As String 'begin with "!STANDARD!" if predefined operator otherwise set function name
End Type

Public Enum NODETYPE
    ''' TODO: boolean
    nt_type = 1
    nt_const
    nt_variable
    nt_arglist
    nt_binary
    nt_statementlist
    nt_makestatement
    nt_callstatement
    nt_dim
    nt_function
    nt_prototype
End Enum

Public Enum METHODTYPE
    mt_function = &H1
    mt_sub = &H2
    mt_public = &H10
    mt_private = &H20
End Enum

Public Enum ARGUMENTTYPE
    at_byref
    at_byval
End Enum

Public Enum CGSTEP
    cg_type = &H1
    cg_const = &H2
    cg_def = &H4
    cg_all = &H10000000
    cg_valid_mask = &H10000007
End Enum

Public Enum STANDARD_METHOD
    sm_retval = 0 'do nothing but return source value
    sm_add '(i,i)i
    sm_sub '(i,i)i
    sm_mul '(i,i)i
    sm_div '(i,i)i
    sm_mod '(i,i)i
    sm_fadd '(f,f)f
    sm_fsub '(f,f)f
    sm_fmul '(f,f)f
    sm_fdiv '(f,f)f
    sm_udiv '(u,u)u
    sm_umod '(u,u)u
    sm_neg '(i)i
    sm_fneg '(f)f
    sm_pow ''' TODO:
    sm_and '(i,i)i
    sm_or '(i,i)i
    sm_xor '(i,i)i
    sm_not '(i)i
    sm_shl '(i,i)i
    sm_shr '(i,i)i
    sm_ushr '(u,u)u
    sm_rol ''' TODO:
    sm_ror ''' TODO:
    sm_land ''' TODO:
    sm_lor ''' TODO:
End Enum
