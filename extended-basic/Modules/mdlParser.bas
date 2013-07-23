Attribute VB_Name = "mdlParser"
Option Explicit

'****************************************
' Parser
' code parser support
'****************************************

#Const PARSER_DEBUG_TRACE = False
Dim m_ParserDebugTraceLevel As Long

Public Sub SyntaxInit()
'    ReDim BinopPrecedence(0) As Long
'    SetPre token_land, 0  '&&
'    SetPre token_lor, 0 '||
'    SetPre token_plus, 10 '+
'    SetPre token_minus, 10 '-
'    SetPre token_mul, 20 '*
'    SetPre token_div, 20 '/
'    SetPre token_mod, 20 '%
'    SetPre token_power, 30 '^
'    SetPre token_and, 40 '&
'    SetPre token_or, 40 '|
'    SetPre token_xor, 40 '^|
'    SetPre token_shl, 50 '<<
'    SetPre token_shr, 50 '>>
'    SetPre token_rol, 50 '<<<
'    SetPre token_ror, 50 '>>>
End Sub

Public Function ParseBinopRHS(ByVal Src As SourceFile, ByVal MinPrecedence As Long, ByVal LHS As IASTNode) As IASTNode
    Dim TokenPrecedence As Long, NextPrecedence As Long
    Dim BinOp As Long
    Dim RHS As IASTNode
    Dim Node As BinaryNode
    Dim Loc As Location
    Loc.File = Src.GetFileName
    Loc.Line = CurTok.Line
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceEnter "ParseBinopRHS", Loc
#End If
    Do
        TokenPrecedence = GetTokenPrecedence
        If TokenPrecedence < MinPrecedence Then
            Set ParseBinopRHS = LHS
            GoTo Final
        End If
        BinOp = CurTok.TokenType
        GetNextToken Src  'eat the binop
        If UnContinuableError Then GoTo Final
        Set RHS = ParseTerm(Src)
        If RHS Is Nothing Then GoTo Final
        NextPrecedence = GetTokenPrecedence
        If TokenPrecedence < NextPrecedence Then
            Set RHS = ParseBinopRHS(Src, TokenPrecedence + 1, RHS)
            If RHS Is Nothing Then GoTo Final
        End If
        Set Node = New BinaryNode
        Node.Create Loc, BinOp, LHS, RHS
        Set LHS = Node
    Loop
Final:
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceExit
#End If
End Function

Public Function ParseCall(ByVal Src As SourceFile, Callee As String) As IASTNode
    Dim Node As New CallStatementNode, Arg As New ArgListNode, p As IASTNode
    Dim Loc As Location
    Loc.File = Src.GetFileName
    Loc.Line = CurTok.Line
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceEnter "ParseCall", Loc
#End If
    Do
        GetNextToken Src 'eat '( 'or ','
        If UnContinuableError Then GoTo Final
        If CurTok.TokenType = token_rbraket Then Exit Do
        Set p = ParseExpression(Src)
        If p Is Nothing Then GoTo Final
        Arg.AddSubNode p
        If CurTok.TokenType = token_rbraket Then Exit Do
        If CurTok.TokenType <> token_comma Then
            SyntaxError Src, "expected ')' or ','"
            If UnContinuableError Then GoTo Final
        End If
    Loop
    GetNextToken Src 'eat ')'
    If UnContinuableError Then GoTo Final
    Node.Create Loc, Callee, Arg
    Set ParseCall = Node
Final:
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceExit
#End If
End Function

Public Function ParseConst(ByVal Src As SourceFile) As IASTNode
    Dim Node As New ConstNode
    Dim T As TypeNode
    Dim Loc As Location
    Loc.File = Src.GetFileName
    Loc.Line = CurTok.Line
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceEnter "ParseConst", Loc
#End If
    Select Case CurTok.TokenType
    Case token_string
        Set T = TypeByName("String")
        If T Is Nothing Then
            FatalError "undefined basic type 'String'"
            GoTo Final
        End If
    Case token_integer
        If Val(CurTok.Identifier) > CDbl(4294967296#) Then
            Set T = TypeByName("LongLong")
            If T Is Nothing Then
                FatalError "undefined basic type 'LongLong'"
                GoTo Final
            End If
        Else
            Set T = TypeByName("Long")
            If T Is Nothing Then
                FatalError "undefined basic type 'Long'"
                GoTo Final
            End If
        End If
    Case token_float
        Set T = TypeByName("Single")
        If T Is Nothing Then
            FatalError "undefined basic type 'Single'"
            GoTo Final
        End If
    End Select
    Node.Create Loc, CurTok, T
    GetNextToken Src 'eat the const
    If UnContinuableError Then GoTo Final
    Set ParseConst = Node
Final:
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceExit
#End If
End Function

Public Function ParseDim(ByVal Src As SourceFile, Optional ByVal SymTable As Dictionary) As IASTNode
    'Note: CurTok is 'Dim' if it's a dim and ')' if it's a prototype
    ''' TODO: comma
    Dim Node As New DimNode
    Dim Name As String
    Dim Loc As Location
    Loc.File = Src.GetFileName
    Loc.Line = CurTok.Line
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceEnter "ParseDim", Loc
#End If
    GetNextToken Src 'eat 'Dim'
    If UnContinuableError Then GoTo Final
    If CurTok.TokenType <> token_identifier Then
        SyntaxError Src, "expected identifier"
        If UnContinuableError Then GoTo Final
    End If
    Name = CurTok.Identifier ''' TODO: variable name processing
    GetNextToken Src 'eat identifier
    If UnContinuableError Then GoTo Final
    If CurTok.TokenType <> keyword_as Then
        SyntaxError Src, "expected 'As'"
        If UnContinuableError Then GoTo Final
    End If
    GetNextToken Src 'eat 'As'
    If UnContinuableError Then GoTo Final
    If CurTok.TokenType <> token_identifier Then
        SyntaxError Src, "expected type identifier"
        If UnContinuableError Then GoTo Final
    End If
    Node.Create Loc, Name, CurTok.Identifier
    GetNextToken Src 'eat type identifier
    If UnContinuableError Then GoTo Final
    ''' TODO: parse multiple dim
    If SymTable Is Nothing Then
        Set SymTable = Src.SymTable
    End If
    If SymTable.Exists(Name) Then
        SyntaxError Src, "'" & Name & "' already exists"
        If UnContinuableError Then GoTo Final
    End If
    SymTable.Add Name, Node
    Set ParseDim = Node
Final:
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceExit
#End If
End Function

Public Function ParseExpression(ByVal Src As SourceFile) As IASTNode
    Select Case CurTok.TokenType
    Case token_string
        Set ParseExpression = ParseConst(Src)
    Case Else
        ''' TODO: parse right-handed unary
        Dim LHS As IASTNode
        Set LHS = ParseTerm(Src)
        If LHS Is Nothing Then Exit Function
        Set ParseExpression = ParseBinopRHS(Src, 0, LHS)
    End Select
End Function

Public Function ParseFile(ByVal Src As SourceFile) As Long
    Dim Node As IASTNode
    Dim Attr As METHODTYPE
    Do
        Do
            GetNextToken Src 'eat CRLF
            If UnContinuableError Then Exit Function
        Loop While CurTok.TokenType = token_crlf
        If CurTok.TokenType = token_eof Then
            Exit Do
        End If
        Attr = mt_public 'default value
        Select Case CurTok.TokenType
        Case keyword_public
            Attr = mt_public
            GetNextToken Src 'eat 'Public'
            If UnContinuableError Then Exit Function
        Case keyword_private
            Attr = mt_private
            GetNextToken Src 'eat 'Private'
            If UnContinuableError Then Exit Function
        End Select
        Select Case CurTok.TokenType
        Case token_eof
            Exit Do
        Case keyword_function
            Attr = Attr Or mt_function
            Set Node = ParseFunction(Src, Attr)
        Case keyword_sub
            Attr = Attr Or mt_sub
            Set Node = ParseFunction(Src, Attr)
        ''' TODO:
        Case Else
            SyntaxError Src, "unexpected '" & GetTokenName(CurTok.TokenType) & "'"
            If UnContinuableError Then Exit Function
        End Select
        If Node Is Nothing Then Exit Function
    Loop
    ParseFile = 1
End Function

Public Function ParseFunction(ByVal Src As SourceFile, ByVal Attributes As METHODTYPE) As IASTNode
    'Note: when this function is called, CurTok points to 'Function' or 'Sub' or other
    Dim Node As New FunctionNode
    Dim Name As String, Proto As PrototypeNode, Ret As DimNode
    Dim Statement As StatementListNode
    Dim EndToken As TOKEN_ENUM
    Dim Loc As Location
    Loc.File = Src.GetFileName
    Loc.Line = CurTok.Line
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceEnter "ParseFunction", Loc
#End If
    EndToken = CurTok.TokenType
    GetNextToken Src 'eat 'Function' or 'Sub' or other
    If UnContinuableError Then GoTo Final
    If CurTok.TokenType <> token_identifier Then
        SyntaxError Src, "expected identifier"
        If UnContinuableError Then GoTo Final
    End If
    Name = CurTok.Identifier
    ''' TODO: other keywords
    GetNextToken Src 'eat identifier
    If UnContinuableError Then GoTo Final
    Set Proto = ParsePrototype(Src)
    If Proto Is Nothing Then GoTo Final
    If Attributes And mt_function Then
        If CurTok.TokenType = keyword_as Then
            GetNextToken Src 'eat 'As'
            If UnContinuableError Then GoTo Final
            Set Ret = New DimNode
            Ret.Create Loc, "Function", CurTok.Identifier
            GetNextToken Src 'eat type identifier
            If UnContinuableError Then GoTo Final
        Else
            SyntaxError Src, "expected 'As'"
            If UnContinuableError Then GoTo Final
        End If
    End If
    If CurTok.TokenType <> token_crlf Then
        SyntaxError Src, "expected crlf"
        If UnContinuableError Then GoTo Final
    End If
    GetNextToken Src 'eat crlf
    If UnContinuableError Then GoTo Final
    Set Statement = ParseStatementList(Src, EndToken, Node.SymTable)
    Node.Create Loc, Name, Attributes, Proto, Ret, Statement
    Src.SymTable.Add Name, Node
    Set ParseFunction = Node
Final:
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceExit
#End If
End Function

Public Function ParseMake(ByVal Src As SourceFile, ByVal Variable As VariableNode) As IASTNode
    ''' TODO: operator extension
    Dim Node As New MakeStatementNode
    Dim RHS As IASTNode
    Dim TmpNode As BinaryNode
    Dim mt As TOKEN_ENUM
    Dim Loc As Location
    Loc.File = Src.GetFileName
    Loc.Line = CurTok.Line
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceEnter "ParseMake", Loc
#End If
    mt = CurTok.TokenType
    GetNextToken Src 'eat the make token
    If UnContinuableError Then GoTo Final
    Set RHS = ParseExpression(Src)
    If RHS Is Nothing Then GoTo Final
    '''TODO:
    'If mt = token_make Then
    Node.Create Loc, Variable, RHS
    Set ParseMake = Node
    'Else
    'End If
Final:
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceExit
#End If
End Function

Public Function ParseParen(ByVal Src As SourceFile) As IASTNode
    Dim Node As IASTNode
    GetNextToken Src 'eat '('
    If UnContinuableError Then Exit Function
    Set Node = ParseExpression(Src)
    If CurTok.TokenType <> token_rbraket Then
        SyntaxError Src, "expected ')'"
        If UnContinuableError Then Exit Function
    End If
    GetNextToken Src 'eat ')'
    If UnContinuableError Then Exit Function
    Set ParseParen = Node
End Function

Public Function ParsePrototype(ByVal Src As SourceFile) As IASTNode
    Dim Node As New PrototypeNode
    Dim ArgName As DimNode, ArgAttr As ARGUMENTTYPE
    Dim s As String
    Dim ArgIndex As Long
    Dim Loc As Location
    Loc.File = Src.GetFileName
    Loc.Line = CurTok.Line
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceEnter "ParsePrototype", Loc
#End If
    Node.Create Loc
    GoTo PrototypeStart
    Do
        If CurTok.TokenType = token_eof Then
            SyntaxError Src, "incomplete prototype", False
             If UnContinuableError Then GoTo Final
        ElseIf CurTok.TokenType = token_comma Then
            'continue
        ElseIf CurTok.TokenType = token_rbraket Then
            Exit Do
        Else
            SyntaxError Src, "expected ')'"
            If UnContinuableError Then GoTo Final
        End If
PrototypeStart:
        GetNextToken Src 'eat '(' or ','
        If UnContinuableError Then GoTo Final 'for case 'xxx()' without any params
        If CurTok.TokenType = token_rbraket Then
            Exit Do
        End If
        ArgAttr = 0
        Select Case CurTok.TokenType
        Case keyword_byval
            ArgAttr = ArgAttr Or at_byval
            GetNextToken Src 'eat 'ByVal'
            If UnContinuableError Then GoTo Final
        Case keyword_byref
            ArgAttr = ArgAttr Or at_byref
            GetNextToken Src 'eat 'ByRef'
            If UnContinuableError Then GoTo Final
        End Select
        If CurTok.TokenType <> token_identifier Then
            SyntaxError Src, "expected identifier"
            If UnContinuableError Then GoTo Final
        End If
        s = CurTok.Identifier
        GetNextToken Src 'eat identifier
        If UnContinuableError Then GoTo Final
        If CurTok.TokenType <> keyword_as Then
            SyntaxError Src, "expected 'As'"
            If UnContinuableError Then GoTo Final
        End If
        GetNextToken Src 'eat 'As'
        If UnContinuableError Then GoTo Final
        If CurTok.TokenType <> token_identifier Then
            SyntaxError Src, "expected type identifier"
            If UnContinuableError Then GoTo Final
        End If
        Set ArgName = New DimNode
        Loc.File = Src.GetFileName
        Loc.Line = CurTok.Line
        ArgName.Create Loc, s, CurTok.Identifier, ArgIndex
        ArgIndex = ArgIndex + 1
        Node.AddParam ArgName, ArgAttr
        GetNextToken Src 'eat type identifier
        If UnContinuableError Then GoTo Final
    Loop
    GetNextToken Src 'eat ')'
    If UnContinuableError Then GoTo Final
    Set ParsePrototype = Node
Final:
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceExit
#End If
End Function

Public Function ParseStatementList(ByVal Src As SourceFile, ByVal EndToken2 As TOKEN_ENUM, ByVal SymTable As Dictionary) As IASTNode
    'Note: EndToken2 indicates which token with "End" is the end of statement list, such as keyword_function for 'End Function'
    Dim Node As New StatementListNode
    Dim v As VariableNode
    Dim SubNode As IASTNode
    Dim Loc As Location
    Loc.File = Src.GetFileName
    Loc.Line = CurTok.Line
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceEnter "ParseStatementList", Loc
#End If
    Node.Create Loc
    Do
        ''' TODO: other separators
        Do While CurTok.TokenType = token_crlf
            GetNextToken Src 'eat CRLF
            If UnContinuableError Then GoTo Final
        Loop
        If CurTok.TokenType = token_eof Then
            Exit Do
        End If
        If CurTok.TokenType = keyword_function Then
            If EndToken2 = keyword_function Then 'in function
                CurTok.TokenType = token_identifier
            End If
        End If
        Select Case CurTok.TokenType
        Case token_identifier
            Set v = ParseVariable(Src)
            If v Is Nothing Then GoTo Final
            If OperatorByName(CurTok.Identifier).Flags And of_make Then 'make statement
                Set SubNode = ParseMake(Src, v)
            End If
        Case keyword_dim
            Set SubNode = ParseDim(Src, SymTable)
        Case token_eof
            Exit Do
        Case keyword_end
            GetNextToken Src 'eat 'End'
            If CurTok.TokenType <> EndToken2 Then
                SyntaxError Src, "expected 'End " & GetTokenName(EndToken2) & "'"
                If UnContinuableError Then GoTo Final
            End If
            Exit Do
        Case Else
            SyntaxError Src, "invalid or currently unsupported '" & GetTokenName(CurTok.TokenType) & "'"
            If UnContinuableError Then GoTo Final
        End Select
        If SubNode Is Nothing Then GoTo Final
        Node.AddSubNode SubNode
    Loop
    Set ParseStatementList = Node
Final:
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceExit
#End If
End Function

Public Function ParseTerm(ByVal Src As SourceFile) As IASTNode
    Select Case CurTok.TokenType
    Case token_identifier
        Set ParseTerm = ParseVariable(Src)
    Case token_integer, token_float
        Set ParseTerm = ParseConst(Src)
    Case token_lbraket
        Set ParseTerm = ParseParen(Src)
    Case keyword_function
        Set ParseTerm = ParseVariable(Src)
    Case Else
        SyntaxError Src, "unexpected " & GetTokenName(CurTok.TokenType)
    End Select
End Function

Public Function ParseVariable(ByVal Src As SourceFile) As IASTNode
    Dim Node As New VariableNode
    Dim Loc As Location
    Loc.File = Src.GetFileName
    Loc.Line = CurTok.Line
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceEnter "ParseVariable", Loc
#End If
    Node.Create Loc, CurTok
    GetNextToken Src                                        'eat the variable
    If UnContinuableError Then GoTo Final
    If CurTok.TokenType = token_lbraket Then
        Set ParseVariable = ParseCall(Src, Node.Name)
    Else
        Set ParseVariable = Node
    End If
Final:
#If PARSER_DEBUG_TRACE Then
    ParserDebugTraceExit
#End If
End Function

Private Sub SyntaxError(ByVal Src As SourceFile, s As String, Optional ByVal Continuable As Boolean = True)
#If PARSER_DEBUG_TRACE Then
    PrintLine String$(m_ParserDebugTraceLevel * 2, "-") & "CurTok: " & GetTokenName(CurTok.TokenType) & vbTab & CurTok.Identifier
#End If
    PrintError Src.GetFileName, CurTok.Line, s
    If Not Continuable Then ErrorBreak
End Sub

'Private Sub SetPre(ByVal BinOp As Long, ByVal Precedence As Long)
'    If BinOp < 0 Then Exit Sub
'    If UBound(BinopPrecedence) < BinOp Then
'        ReDim Preserve BinopPrecedence(BinOp) As Long
'    End If
'    BinopPrecedence(BinOp) = Precedence
'End Sub
'
'Private Function GetTokenPrecedence() As Long
'    If (GetTokenFlag(CurTok.TokenType) And TK_BINOP) = 0 Then
'        GetTokenPrecedence = -1
'        Exit Function
'    End If
'    GetTokenPrecedence = BinopPrecedence(CurTok.TokenType)
'End Function

Private Function GetTokenPrecedence() As Long
    Dim Op As Operator
    Set Op = OperatorByName(CurTok.Identifier)
    If Not Op Is Nothing Then
        If Op.Flags And of_binop Then
            GetTokenPrecedence = Op.Precedence
        Else
            GetTokenPrecedence = -1
        End If
    Else
        GetTokenPrecedence = -1
    End If
End Function

#If PARSER_DEBUG_TRACE Then
Private Sub ParserDebugTraceEnter(ParserName As String, Loc As Location)
    m_ParserDebugTraceLevel = m_ParserDebugTraceLevel + 1
    PrintLine String$(m_ParserDebugTraceLevel * 2, "-") & ParserName & ": " & Loc.File & "(" & Loc.Line & ")"
End Sub

Private Sub ParserDebugTraceExit()
    m_ParserDebugTraceLevel = m_ParserDebugTraceLevel - 1
End Sub
#End If
