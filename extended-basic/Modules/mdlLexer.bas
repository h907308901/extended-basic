Attribute VB_Name = "mdlLexer"
Option Explicit

'****************************************
' Lexer
' keyword, symbol, token support
'****************************************

#Const LEXER_DEBUG_TRACE = False

Public CurTok As TOKEN_TYPE
Public KeyWords() As String '0-based, but nothing in 0th
Public Symbols() As String '0-based, but nothing in 0th
Public ReservedSymbols() As String '0-based, but nothing in 0th
''' Note: a reserved symbol can ONLY contain ONE char and user cannot define any or any symbols containing reserved symbols as operators

Public Sub LexerInit()
    ReDim KeyWords(0) As String
    AddKeyWord "As"
    AddKeyWord "ByRef"
    AddKeyWord "ByVal"
    AddKeyWord "Dim"
    AddKeyWord "End"
    AddKeyWord "Function"
    AddKeyWord "Sub"
    AddKeyWord "Private"
    AddKeyWord "Public"
    ReDim ReservedSymbols(0) As String
    AddReservedSymbol "."
    AddReservedSymbol ","
    AddReservedSymbol "("
    AddReservedSymbol ")"
    'predefined symbols, user may redefine them
    ReDim Symbols(0) As String
    AddSymbol "="
    AddSymbol "+="
    AddSymbol "-="
    AddSymbol "*="
    AddSymbol "/="
    AddSymbol "^="
    AddSymbol "&="
    AddSymbol "|="
    AddSymbol "^|="
    AddSymbol "<<="
    AddSymbol ">>="
    AddSymbol "<<<="
    AddSymbol ">>>="
    AddSymbol "~"
    AddSymbol "=="
    AddSymbol "<"
    AddSymbol ">"
    AddSymbol "<="
    AddSymbol ">="
    AddSymbol "<>"
    AddSymbol "&&"
    AddSymbol "||"
    AddSymbol "+"
    AddSymbol "-"
    AddSymbol "*"
    AddSymbol "/"
    AddSymbol "%"
    AddSymbol "^"
    AddSymbol "&"
    AddSymbol "|"
    AddSymbol "^|"
    AddSymbol "<<"
    AddSymbol ">>"
    AddSymbol "<<<"
    AddSymbol ">>>"
    AddSymbol "++"
    AddSymbol "--"
End Sub

Public Sub GetNextToken(ByVal Src As SourceFile)
    Dim Char As VBCharType
    Dim i As Long
    CurTok.Identifier = ""
    CurTok.IsError = 0
TokenStart:
    Do
        Char = Src.GetChar
        If Char = ["\0"] Then Char = [" "]
    Loop While Char = ["\t"] Or Char = [" "] Or Char = ["\r"]
    CurTok.Line = Src.GetLine
    Select Case Char
    Case [EOF]                                              'end of file
EOFStart:
        CurTok.TokenType = token_eof
    Case ["\n"]                                             'crlf
        CurTok.TokenType = token_crlf
    Case ["'"]                                              'note
        Do
            Char = Src.GetChar
            If Char = [EOF] Then GoTo EOFStart
        Loop Until Char = ["\r"] Or Char = ["\n"]
        GoTo TokenStart
    Case ["""] 'string
        CurTok.TokenType = token_string
        GoTo StrStart
        Do
            CurTok.Identifier = CurTok.Identifier & ChrW$(Char)
StrStart:
            Char = Src.GetChar
            Select Case Char
            Case [EOF], ["\r"], ["\n"]
                LexerError Src, "expected '"                '"
            End Select
            'case of quotation mark in string
            If Char = ["""] Then
                Char = Src.GetChar
                If Char <> ["""] Then
                    Src.UnGetChar
                    Exit Do
                End If
            End If
        Loop
    Case ["0"] To ["9"]                                     'numeric
        CurTok.TokenType = token_integer
        Do
            CurTok.Identifier = CurTok.Identifier & ChrW$(Char)
            Char = Src.GetChar
            Select Case Char
            Case ["0"] To ["9"]                             'continue
            Case ["."]
                If CurTok.TokenType = token_float Then
                    LexerError Src, "unexpected '.'"
                    Exit Do
                Else
                    CurTok.TokenType = token_float
                End If
            Case Else
                Src.UnGetChar
                Exit Do
            End Select
        Loop
    Case ["_"], ["a"] To ["z"], ["aa"] To ["zz"]            'identifier
        CurTok.TokenType = token_identifier
        Do
            CurTok.Identifier = CurTok.Identifier & ChrW$(Char)
            Char = Src.GetChar
            Select Case Char
            Case ["_"], ["a"] To ["z"], ["aa"] To ["zz"], ["0"] To ["9"] 'continue
            Case Else
                Src.UnGetChar
                Exit Do
            End Select
        Loop
        For i = 1 To UBound(KeyWords)
            If CurTok.Identifier = KeyWords(i) Then
                CurTok.TokenType = token_keyword_start + i
            End If
        Next
    Case Else
        If IsSymbol(Char) Then
            If IsReservedSymbol(Char) Then
                CurTok.Identifier = ChrW$(Char)
            Else
                Do
                    CurTok.Identifier = CurTok.Identifier & ChrW$(Char)
                    Char = Src.GetChar
                    If IsReservedSymbol(Char) Or Not IsSymbol(Char) Then
                        Src.UnGetChar
                        Exit Do
                    End If
                Loop
            End If
            CurTok.TokenType = TokenBySymbol(CurTok.Identifier)
        End If
    End Select
'    Case ["."]                                              'dot
'        CurTok.TokenType = token_dot
'    Case [","]                                              'comma
'        CurTok.TokenType = token_comma
'    Case ["("]                                              'lbraket
'        CurTok.TokenType = token_lbraket
'    Case [")"]                                              'rbraket
'        CurTok.TokenType = token_rbraket
'    Case ["="]
'        Char = Src.GetChar
'        If Char = ["="] Then
'            CurTok.TokenType = token_equal                  '==
'        Else
'            Src.UnGetChar
'            CurTok.TokenType = token_make                   '=
'        End If
'    Case ["<"]
'        Char = Src.GetChar
'        Select Case Char
'        Case ["="]
'            CurTok.TokenType = token_le                     '<=
'        Case ["<"]
'            Char = Src.GetChar
'            Select Case Char
'            Case ["="]
'                CurTok.TokenType = token_mshl               '<<=
'            Case ["<"]
'                Char = Src.GetChar
'                Select Case Char
'                Case ["="]
'                    CurTok.TokenType = token_mrol           '<<<=
'                Case Else
'                    Src.UnGetChar
'                    CurTok.TokenType = token_rol            '<<<
'                End Select
'            Case Else
'                Src.UnGetChar
'                CurTok.TokenType = token_shl                '<<
'            End Select
'        Case [">"]
'            CurTok.TokenType = token_ne                     '<>
'        Case Else
'            Src.UnGetChar
'            CurTok.TokenType = token_lt                     '<
'        End Select
'    Case [">"]
'        Char = Src.GetChar
'        Select Case Char
'        Case ["="]
'            CurTok.TokenType = token_ge                     '>=
'        Case [">"]
'            Char = Src.GetChar
'            Select Case Char
'            Case ["="]
'                CurTok.TokenType = token_mshr               '>>=
'            Case [">"]
'                Char = Src.GetChar
'                Select Case Char
'                Case ["="]
'                    CurTok.TokenType = token_mror           '>>>=
'                Case Else
'                    Src.UnGetChar
'                    CurTok.TokenType = token_ror            '>>>
'                End Select
'            Case Else
'                Src.UnGetChar
'                CurTok.TokenType = token_shr                '>>
'            End Select
'        Case Else
'            Src.UnGetChar
'            CurTok.TokenType = token_gt                     '>
'        End Select
'    Case ["+"]
'        Char = Src.GetChar
'        Select Case Char
'        Case ["="]
'            CurTok.TokenType = token_mplus                  '+=
'        Case ["+"]
'            CurTok.TokenType = token_plus1                  '++
'        Case Else
'            Src.UnGetChar
'            CurTok.TokenType = token_plus                   '+
'        End Select
'    Case ["-"]
'        Char = Src.GetChar
'        Select Case Char
'        Case ["="]
'            CurTok.TokenType = token_mminus                 '-=
'        Case ["-"]
'            CurTok.TokenType = token_minus1                 '--
'        Case Else
'            Src.UnGetChar
'            CurTok.TokenType = token_minus                  '-
'        End Select
'    Case ["*"]
'        Char = Src.GetChar
'        If Char = ["="] Then
'            CurTok.TokenType = token_mmul                   '*=
'        Else
'            Src.UnGetChar
'            CurTok.TokenType = token_mul                    '*
'        End If
'    Case ["/"]
'        Char = Src.GetChar
'        If Char = ["="] Then
'            CurTok.TokenType = token_mdiv                   '/=
'        Else
'            Src.UnGetChar
'            CurTok.TokenType = token_div                    '/
'        End If
'    Case ["%"]
'        CurTok.TokenType = token_mod                        '%
'    Case ["^"]
'        Char = Src.GetChar
'        Select Case Char
'        Case ["="]
'            CurTok.TokenType = token_mpower                 '^=
'        Case ["|"]
'            Char = Src.GetChar
'            If Char = ["="] Then
'                CurTok.TokenType = token_mxor               '^|=
'            Else
'                Src.UnGetChar
'                CurTok.TokenType = token_xor                '^|
'            End If
'        Case Else
'            Src.UnGetChar
'            CurTok.TokenType = token_power                  '^
'        End Select
'    Case ["&"]
'        Char = Src.GetChar
'        Select Case Char
'        Case ["="]
'            CurTok.TokenType = token_mand                   '&=
'        Case ["&"]
'            CurTok.TokenType = token_land                   '&&
'        Case Else
'            Src.UnGetChar
'            CurTok.TokenType = token_and                    '&
'        End Select
'    Case ["|"]
'        Char = Src.GetChar
'        Select Case Char
'        Case ["="]
'            CurTok.TokenType = token_mor                    '|=
'        Case ["|"]
'            CurTok.TokenType = token_lor                    '||
'        Case Else
'            Src.UnGetChar
'            CurTok.TokenType = token_or                     '|
'        End Select
'    Case ["~"]
'        CurTok.TokenType = token_not
'    Case Else
'        LexerError Src, "unexpected '" & ChrW$(Char) & "'"
'    End Select
    'for debug
#If LEXER_DEBUG_TRACE Then
    PrintLine "GetNextToken: " & GetTokenName(CurTok.TokenType) & vbTab & CurTok.Identifier
#End If
End Sub

Public Function GetTokenName(ByVal Token As Long) As String
    Select Case Token
    Case token_eof: GetTokenName = "<EOF>"
    Case token_crlf: GetTokenName = "<CRLF>"
    Case token_identifier: GetTokenName = "<ID>"
    Case token_integer: GetTokenName = "<INT>"
    Case token_float: GetTokenName = "<FLOAT>"
    Case token_string: GetTokenName = "<STR>"
    Case Is > token_symbol_start: GetTokenName = Symbols(Token - token_symbol_start)
    Case Is > token_keyword_start: GetTokenName = KeyWords(Token - token_keyword_start)
    Case Is > token_reserved_symbol_start: GetTokenName = ReservedSymbols(Token - token_reserved_symbol_start)
    End Select
End Function

Public Function TokenBySymbol(Symbol As String) As TOKEN_ENUM
    Dim i As Long
    TokenBySymbol = token_unknown_symbol
    For i = 1 To UBound(ReservedSymbols)
        If Symbol = ReservedSymbols(i) Then
            TokenBySymbol = token_reserved_symbol_start + i
            Exit Function
        End If
    Next
    For i = 1 To UBound(Symbols)
        If Symbol = Symbols(i) Then
            TokenBySymbol = token_symbol_start + i
            Exit Function
        End If
    Next
End Function

Private Function IsSymbol(ByVal Char As VBCharType) As Boolean
    Select Case Char
    Case ["!"] To ["/"], ["col"] To ["@"], ["lll"] To ["`"], ["{"] To ["~"]
        IsSymbol = True
    End Select
End Function

Private Function IsReservedSymbol(ByVal Char As VBCharType) As Boolean
    Dim i As Long
    For i = 1 To UBound(ReservedSymbols)
        If ChrW$(Char) = ReservedSymbols(i) Then
            IsReservedSymbol = True
            Exit Function
        End If
    Next
End Function

Private Sub AddKeyWord(s As String)
    'Note: keywords must be defined in TOKEN_ENUM first and the order must be right
    ReDim Preserve KeyWords(UBound(KeyWords) + 1) As String
    KeyWords(UBound(KeyWords)) = s
End Sub

Private Sub AddSymbol(s As String)
    ReDim Preserve Symbols(UBound(Symbols) + 1) As String
    Symbols(UBound(Symbols)) = s
End Sub

Private Sub AddReservedSymbol(s As String)
    'Note: reserved symbols must be defined in TOKEN_ENUM first and the order must be right
    ReDim Preserve ReservedSymbols(UBound(ReservedSymbols) + 1) As String
    ReservedSymbols(UBound(ReservedSymbols)) = s
End Sub

Private Sub LexerError(ByVal Src As SourceFile, s As String, Optional ByVal Continuable As Boolean = True)
    PrintError Src.GetFileName, CurTok.Line, s
    CurTok.IsError = 1
    If Not Continuable Then ErrorBreak
End Sub
