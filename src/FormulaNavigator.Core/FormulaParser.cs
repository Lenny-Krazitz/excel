using System;
using System.Collections.Generic;
using System.Globalization;

namespace FormulaNavigator.Core
{
    public static class FormulaParser
    {
        public static FormulaParseResult Parse(string formula)
        {
            if (formula == null)
                return new FormulaParseResult(null, new List<string> { "Formula is null." });
            Parser parser = new Parser(formula);
            return parser.Parse();
        }

        private enum TokenKind
        {
            End, Number, String, Identifier, Reference, Error, OpenParen, CloseParen,
            Comma, Semicolon, OpenBrace, CloseBrace, Colon, Operator, Percent, Hash, At,
            StructuredReference, MalformedStructuredReference, Unknown
        }

        private sealed class Token
        {
            public TokenKind Kind;
            public int Start;
            public int Length;
            public bool LeadingWhitespace;
            public string Text;
        }

        private sealed class Lexer
        {
            private readonly string _source;
            private int _position;

            public Lexer(string source) { _source = source; }

            public Token Next()
            {
                bool whitespace = false;
                while (_position < _source.Length && Char.IsWhiteSpace(_source[_position]))
                {
                    whitespace = true;
                    _position++;
                }
                if (_position >= _source.Length) return New(TokenKind.End, _position, 0, whitespace);
                int start = _position;
                char c = _source[_position++];
                if (c == '"')
                {
                    bool closed = false;
                    while (_position < _source.Length)
                    {
                        if (_source[_position] == '"')
                        {
                            _position++;
                            if (_position < _source.Length && _source[_position] == '"') { _position++; continue; }
                            closed = true; break;
                        }
                        _position++;
                    }
                    return New(closed ? TokenKind.String : TokenKind.Unknown, start, _position - start, whitespace);
                }
                if (c == '\'')
                {
                    bool closed = false;
                    while (_position < _source.Length)
                    {
                        if (_source[_position] == '\'')
                        {
                            _position++;
                            if (_position < _source.Length && _source[_position] == '\'') { _position++; continue; }
                            closed = true; break;
                        }
                        _position++;
                    }
                    if (closed && _position < _source.Length && _source[_position] == '!')
                    {
                        _position++;
                        ReadReferenceTail();
                        return ReferenceOrUnknown(start, whitespace);
                    }
                    return New(TokenKind.Unknown, start, _position - start, whitespace);
                }
                if (c == '[')
                {
                    if (IsStructuredStart(_position))
                    {
                        return New(ReadStructuredReference() ? TokenKind.StructuredReference : TokenKind.MalformedStructuredReference,
                            start, _position - start, whitespace);
                    }
                    int close = _source.IndexOf(']', _position);
                    int bang = close < 0 ? -1 : _source.IndexOf('!', close + 1);
                    if (bang > close && IsUnquotedQualifierTail(close + 1, bang))
                    {
                        _position = bang + 1; ReadReferenceTail(); return ReferenceOrUnknown(start, whitespace);
                    }
                    return New(TokenKind.Unknown, start, 1, whitespace);
                }
                if (Char.IsDigit(c) || (c == '.' && _position < _source.Length && Char.IsDigit(_source[_position])))
                {
                    while (_position < _source.Length && Char.IsDigit(_source[_position])) _position++;
                    if (_position < _source.Length && _source[_position] == '.')
                    {
                        _position++;
                        while (_position < _source.Length && Char.IsDigit(_source[_position])) _position++;
                    }
                    if (_position < _source.Length && (_source[_position] == 'E' || _source[_position] == 'e'))
                    {
                        int exponent = _position++;
                        if (_position < _source.Length && (_source[_position] == '+' || _source[_position] == '-')) _position++;
                        int digits = _position;
                        while (_position < _source.Length && Char.IsDigit(_source[_position])) _position++;
                        if (digits == _position) _position = exponent;
                    }
                    return New(TokenKind.Number, start, _position - start, whitespace);
                }
                if (c == '#')
                {
                    if (_position < _source.Length && Char.IsLetter(_source[_position]))
                    {
                        while (_position < _source.Length && (Char.IsLetterOrDigit(_source[_position])
                            || _source[_position] == '/' || _source[_position] == '!' || _source[_position] == '?'
                            || _source[_position] == '_')) _position++;
                        return New(TokenKind.Error, start, _position - start, whitespace);
                    }
                    return New(TokenKind.Hash, start, 1, whitespace);
                }
                if (c == '$')
                {
                    ReadReferenceTail();
                    string text = _source.Substring(start, _position - start);
                    ReferenceArea ignored;
                    return New(ReferenceParser.TryParse(text, null, null, out ignored) ? TokenKind.Reference : TokenKind.Unknown,
                        start, _position - start, whitespace);
                }
                if (Char.IsLetter(c) || c == '_' || c == '\\')
                {
                    while (_position < _source.Length && IsNameCharacter(_source[_position])) _position++;
                    if (_position < _source.Length && _source[_position] == '$')
                    {
                        _position++;
                        while (_position < _source.Length && Char.IsDigit(_source[_position])) _position++;
                        string referenceText = _source.Substring(start, _position - start);
                        ReferenceArea mixedReference;
                        return New(ReferenceParser.TryParse(referenceText, null, null, out mixedReference)
                            ? TokenKind.Reference : TokenKind.Unknown, start, _position - start, whitespace);
                    }
                    if (_position < _source.Length && _source[_position] == '!')
                    {
                        _position++;
                        ReadReferenceTail();
                        return ReferenceOrUnknown(start, whitespace);
                    }
                    if (_position < _source.Length && _source[_position] == '(')
                        return New(TokenKind.Identifier, start, _position - start, whitespace);
                    if (_position < _source.Length && _source[_position] == '[')
                    {
                        _position++; // Match the bare '[' path: the opening bracket is already consumed.
                        return New(ReadStructuredReference() ? TokenKind.StructuredReference : TokenKind.MalformedStructuredReference,
                            start, _position - start, whitespace);
                    }
                    string text = _source.Substring(start, _position - start);
                    ReferenceArea ignored;
                    return New(ContainsDigit(text) && ReferenceParser.TryParse(text, null, null, out ignored) ? TokenKind.Reference : TokenKind.Identifier,
                        start, _position - start, whitespace);
                }
                switch (c)
                {
                    case '(': return New(TokenKind.OpenParen, start, 1, whitespace);
                    case ')': return New(TokenKind.CloseParen, start, 1, whitespace);
                    case ',': return New(TokenKind.Comma, start, 1, whitespace);
                    case ';': return New(TokenKind.Semicolon, start, 1, whitespace);
                    case '{': return New(TokenKind.OpenBrace, start, 1, whitespace);
                    case '}': return New(TokenKind.CloseBrace, start, 1, whitespace);
                    case ':': return New(TokenKind.Colon, start, 1, whitespace);
                    case '%': return New(TokenKind.Percent, start, 1, whitespace);
                    case '@': return New(TokenKind.At, start, 1, whitespace);
                    case '+': case '-': case '*': case '/': case '^': case '&': case '=': case '<': case '>':
                        if ((c == '<' || c == '>') && _position < _source.Length &&
                            (_source[_position] == '=' || (c == '<' && _source[_position] == '>'))) _position++;
                        return New(TokenKind.Operator, start, _position - start, whitespace);
                    default: return New(TokenKind.Unknown, start, 1, whitespace);
                }
            }

            private void ReadReferenceTail()
            {
                if (_position < _source.Length && _source[_position] == '$') _position++;
                while (_position < _source.Length && Char.IsLetter(_source[_position])) _position++;
                if (_position < _source.Length && _source[_position] == '$') _position++;
                while (_position < _source.Length && Char.IsDigit(_source[_position])) _position++;
            }

            private static bool IsNameCharacter(char c)
            {
                return Char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\\';
            }

            private bool IsStructuredStart(int contentStart)
            {
                return contentStart < _source.Length && (_source[contentStart] == '@' || _source[contentStart] == '#'
                    || _source[contentStart] == '[');
            }

            private bool IsUnquotedQualifierTail(int start, int end)
            {
                if (start == end) return false;
                for (int i = start; i < end; i++)
                {
                    char current = _source[i];
                    if (!Char.IsLetterOrDigit(current) && current != '_' && current != '.') return false;
                }
                return true;
            }

            // Balanced bracket scanning deliberately preserves all content, including header whitespace.
            // Excel's nested bracket syntax (for item specifiers and column ranges) is accepted without
            // attempting to interpret table semantics in the formula layer.
            // The caller has consumed the first opening bracket.
            private bool ReadStructuredReference()
            {
                int depth = 1;
                while (_position < _source.Length)
                {
                    char current = _source[_position++];
                    if (current == '[') depth++;
                    else if (current == ']')
                    {
                        depth--;
                        if (depth == 0) return true;
                        if (depth < 0) return false;
                    }
                }
                return false;
            }

            private static bool ContainsDigit(string value)
            {
                for (int i = 0; i < value.Length; i++) if (Char.IsDigit(value[i])) return true;
                return false;
            }

            private Token New(TokenKind kind, int start, int length, bool whitespace)
            {
                return new Token { Kind = kind, Start = start, Length = length, LeadingWhitespace = whitespace,
                    Text = _source.Substring(start, length) };
            }

            private Token ReferenceOrUnknown(int start, bool whitespace)
            {
                string text = _source.Substring(start, _position - start);
                ReferenceArea ignored;
                return New(ReferenceParser.TryParse(text, null, null, out ignored) ? TokenKind.Reference : TokenKind.Unknown,
                    start, _position - start, whitespace);
            }
        }

        private sealed class Parser
        {
            private const int MaxFormulaLength = 1000000;
            private const int MaxParseDepth = 256;
            private const int MaxTokens = 100000;
            private readonly string _source;
            private readonly Lexer _lexer;
            private readonly List<string> _diagnostics = new List<string>();
            private Token _current;
            private int _commaStops;
            private int _semicolonStops;
            private int _parseDepth;
            private int _tokensSeen;
            private bool _tokenLimitHit;

            public Parser(string source)
            {
                _source = source;
                _lexer = new Lexer(source);
                _current = _lexer.Next();
            }

            public FormulaParseResult Parse()
            {
                if (_source.Length > MaxFormulaLength)
                {
                    Error("Formula exceeds the maximum supported length", MaxFormulaLength);
                    return new FormulaParseResult(null, _diagnostics);
                }
                if (_current.Kind == TokenKind.Operator && _current.Text == "=" && _current.Start == 0) Advance();
                if (_current.Kind == TokenKind.End) Error("Expected an expression", _current.Start);
                FormulaNode root = _current.Kind == TokenKind.End ? null : ParseExpression(0);
                if (_current.Kind != TokenKind.End)
                {
                    Error("Unexpected token '" + _current.Text + "'", _current.Start);
                }
                if (root != null && ExceedsSafeNodeDepth(root))
                {
                    Error("Formula AST exceeds the maximum supported depth", root.Start);
                    root = null;
                }
                return new FormulaParseResult(root, _diagnostics);
            }

            private static bool ExceedsSafeNodeDepth(FormulaNode root)
            {
                Stack<KeyValuePair<FormulaNode, int>> pending = new Stack<KeyValuePair<FormulaNode, int>>();
                pending.Push(new KeyValuePair<FormulaNode, int>(root, 1));
                while (pending.Count > 0)
                {
                    KeyValuePair<FormulaNode, int> current = pending.Pop();
                    if (current.Value > MaxParseDepth) return true;
                    for (int i = 0; i < current.Key.Children.Count; i++)
                    {
                        FormulaNode child = current.Key.Children[i];
                        if (child != null) pending.Push(new KeyValuePair<FormulaNode, int>(child, current.Value + 1));
                    }
                }
                return false;
            }

            private FormulaNode ParseExpression(int minimumPrecedence)
            {
                if (_parseDepth >= MaxParseDepth)
                {
                    Error("Formula nesting exceeds the maximum supported depth", _current.Start);
                    return null;
                }
                _parseDepth++;
                try { return ParseExpressionCore(minimumPrecedence); }
                finally { _parseDepth--; }
            }

            private FormulaNode ParseExpressionCore(int minimumPrecedence)
            {
                FormulaNode left = ParsePrefix();
                if (left == null) return null;
                while (true)
                {
                    if (_current.Kind == TokenKind.Percent || _current.Kind == TokenKind.Hash)
                    {
                        Token suffix = _current; Advance();
                        left = Node(FormulaNodeKind.Unary, left.Start, suffix.Start + suffix.Length - left.Start,
                            suffix.Text, null, left);
                        continue;
                    }
                    string op;
                    int precedence;
                    bool consumed = false;
                    Token operatorToken = _current;
                    if (operatorToken.LeadingWhitespace && IsReferenceExpression(left) && CanStartExpression(operatorToken))
                    {
                        op = " "; precedence = 90;
                    }
                    else if (TryInfix(operatorToken, out op, out precedence))
                    {
                        if ((operatorToken.Kind == TokenKind.Comma && _commaStops > 0)
                            || (operatorToken.Kind == TokenKind.Semicolon && _semicolonStops > 0)) break;
                    }
                    else break;
                    if (precedence < minimumPrecedence) break;
                    if (operatorToken.Kind != TokenKind.End && op != " ") { Advance(); consumed = true; }
                    FormulaNode right = ParseExpression(precedence + 1);
                    if (right == null)
                    {
                        Error("Expected expression after '" + op + "'", operatorToken.Start);
                        return left;
                    }
                    if (op == ":")
                    {
                        left = CoerceRowReference(left);
                        right = CoerceRowReference(right);
                    }
                    left = Node(FormulaNodeKind.Binary, left.Start, right.Start + right.Length - left.Start,
                        op, null, left, right);
                    if (!consumed) { /* intersection is represented by preceding whitespace */ }
                }
                return left;
            }

            private FormulaNode ParsePrefix()
            {
                Token token = _current;
                if (token.Kind == TokenKind.Operator && (token.Text == "+" || token.Text == "-"))
                {
                    Advance();
                    FormulaNode child = ParseExpression(80);
                    if (child == null) { Error("Expected expression after '" + token.Text + "'", token.Start); return null; }
                    return Node(FormulaNodeKind.Unary, token.Start, child.Start + child.Length - token.Start, token.Text, null, child);
                }
                if (token.Kind == TokenKind.At)
                {
                    Advance();
                    FormulaNode child = ParseExpression(80);
                    if (child == null) { Error("Expected expression after '@'", token.Start); return null; }
                    return Node(FormulaNodeKind.Unary, token.Start, child.Start + child.Length - token.Start, "@", null, child);
                }
                if (token.Kind == TokenKind.Number || token.Kind == TokenKind.String || token.Kind == TokenKind.Error)
                {
                    Advance(); return Node(FormulaNodeKind.Literal, token.Start, token.Length, null, null);
                }
                if (token.Kind == TokenKind.Reference)
                {
                    Advance(); return Node(FormulaNodeKind.Reference, token.Start, token.Length, null, null);
                }
                if (token.Kind == TokenKind.StructuredReference)
                {
                    Advance(); return Node(FormulaNodeKind.Name, token.Start, token.Length, null, token.Text);
                }
                if (token.Kind == TokenKind.MalformedStructuredReference)
                {
                    Error("Malformed structured reference", token.Start);
                    Advance(); return Node(FormulaNodeKind.Name, token.Start, token.Length, null, token.Text);
                }
                if (token.Kind == TokenKind.Identifier)
                {
                    Advance();
                    if (String.Equals(token.Text, "TRUE", StringComparison.OrdinalIgnoreCase)
                        || String.Equals(token.Text, "FALSE", StringComparison.OrdinalIgnoreCase))
                        return Node(FormulaNodeKind.Literal, token.Start, token.Length, null, null);
                    if (_current.Kind == TokenKind.OpenParen) return ParseFunction(token);
                    if (_current.Kind == TokenKind.OpenBrace || _current.Kind == TokenKind.Unknown && _current.Text == "[")
                        Error("Structured references are not supported", _current.Start);
                    return Node(FormulaNodeKind.Name, token.Start, token.Length, null, token.Text);
                }
                if (token.Kind == TokenKind.OpenParen)
                {
                    Advance();
                    FormulaNode inner = ParseExpression(0);
                    if (_current.Kind != TokenKind.CloseParen)
                    {
                        Error("Expected ')'", _current.Start);
                        return inner;
                    }
                    Token close = _current; Advance();
                    return Node(FormulaNodeKind.Group, token.Start, close.Start + close.Length - token.Start, null, null, inner);
                }
                if (token.Kind == TokenKind.OpenBrace) return ParseArray();
                if (token.Kind != TokenKind.End) { Error("Expected expression", token.Start); Advance(); }
                return null;
            }

            private FormulaNode ParseFunction(Token name)
            {
                Advance(); // opening parenthesis
                List<FormulaNode> arguments = new List<FormulaNode>();
                _commaStops++;
                if (_current.Kind != TokenKind.CloseParen)
                {
                    while (true)
                    {
                        if (_current.Kind == TokenKind.Comma)
                            arguments.Add(Node(FormulaNodeKind.MissingArgument, _current.Start, 0, null, null));
                        else
                        {
                            FormulaNode argument = ParseExpression(0);
                            if (argument == null) { Error("Expected function argument", _current.Start); break; }
                            arguments.Add(argument);
                        }
                        if (_current.Kind != TokenKind.Comma) break;
                        Token comma = _current; Advance();
                        if (_current.Kind == TokenKind.CloseParen)
                        {
                            arguments.Add(Node(FormulaNodeKind.MissingArgument, comma.Start + comma.Length, 0, null, null));
                            break;
                        }
                    }
                }
                _commaStops--;
                if (_current.Kind != TokenKind.CloseParen)
                {
                    Error("Expected ')' after function arguments", _current.Start);
                    int end = arguments.Count == 0 ? name.Start + name.Length : End(arguments[arguments.Count - 1]);
                    return Node(FormulaNodeKind.Function, name.Start, end - name.Start, null, name.Text, arguments.ToArray());
                }
                Token close = _current; Advance();
                return Node(FormulaNodeKind.Function, name.Start, close.Start + close.Length - name.Start, null, name.Text, arguments.ToArray());
            }

            private FormulaNode ParseArray()
            {
                Token open = _current; Advance();
                List<FormulaNode> items = new List<FormulaNode>();
                _commaStops++; _semicolonStops++;
                while (_current.Kind != TokenKind.CloseBrace && _current.Kind != TokenKind.End)
                {
                    if (_current.Kind == TokenKind.Comma || _current.Kind == TokenKind.Semicolon)
                    {
                        items.Add(Node(FormulaNodeKind.MissingArgument, _current.Start, 0, null, null)); Advance(); continue;
                    }
                    FormulaNode item = ParseExpression(0);
                    if (item == null) break;
                    items.Add(item);
                    if (_current.Kind == TokenKind.Comma || _current.Kind == TokenKind.Semicolon) Advance();
                    else if (_current.Kind != TokenKind.CloseBrace) { Error("Expected array separator", _current.Start); break; }
                }
                _commaStops--; _semicolonStops--;
                if (_current.Kind != TokenKind.CloseBrace)
                {
                    Error("Expected '}'", _current.Start);
                    int end = items.Count == 0 ? open.Start + open.Length : End(items[items.Count - 1]);
                    return Node(FormulaNodeKind.Array, open.Start, end - open.Start, null, null, items.ToArray());
                }
                Token close = _current; Advance();
                return Node(FormulaNodeKind.Array, open.Start, close.Start + close.Length - open.Start, null, null, items.ToArray());
            }

            private static bool TryInfix(Token token, out string op, out int precedence)
            {
                op = token.Text; precedence = 0;
                if (token.Kind == TokenKind.Colon) { precedence = 100; return true; }
                if (token.Kind == TokenKind.Comma || token.Kind == TokenKind.Semicolon) { precedence = 90; return true; }
                if (token.Kind != TokenKind.Operator) return false;
                switch (token.Text)
                {
                    case "^": precedence = 60; return true;
                    case "*": case "/": precedence = 50; return true;
                    case "+": case "-": precedence = 40; return true;
                    case "&": precedence = 30; return true;
                    case "=": case "<": case ">": case "<=": case ">=": case "<>": precedence = 20; return true;
                    default: return false;
                }
            }

            private static bool CanStartExpression(Token token)
            {
                return token.Kind == TokenKind.Number || token.Kind == TokenKind.String || token.Kind == TokenKind.Identifier
                    || token.Kind == TokenKind.Reference || token.Kind == TokenKind.OpenParen || token.Kind == TokenKind.At;
            }

            private static bool IsReferenceExpression(FormulaNode node)
            {
                if (node == null) return false;
                if (node.Kind == FormulaNodeKind.Group && node.Children.Count == 1)
                    return IsReferenceExpression(node.Children[0]);
                return node.Kind == FormulaNodeKind.Reference ||
                    node.Kind == FormulaNodeKind.Binary && (node.Operator == ":" || node.Operator == " " || node.Operator == ",");
            }

            private FormulaNode CoerceRowReference(FormulaNode node)
            {
                if (node == null || (node.Kind != FormulaNodeKind.Literal && node.Kind != FormulaNodeKind.Name)) return node;
                ReferenceArea ignored;
                if (!ReferenceParser.TryParse(node.Text, null, null, out ignored)) return node;
                return Node(FormulaNodeKind.Reference, node.Start, node.Length, null, null);
            }

            private FormulaNode Node(FormulaNodeKind kind, int start, int length, string op, string name, params FormulaNode[] children)
            {
                return new FormulaNode(kind, _source, start, length, op, name, new List<FormulaNode>(children));
            }

            private void Advance()
            {
                if (_tokenLimitHit) { _current = EndToken(); return; }
                _tokensSeen++;
                if (_tokensSeen > MaxTokens)
                {
                    _tokenLimitHit = true;
                    Error("Formula exceeds the maximum supported token count", _current.Start);
                    _current = EndToken();
                    return;
                }
                _current = _lexer.Next();
            }
            private void Error(string message, int position) { _diagnostics.Add(message + " at position " + position.ToString(CultureInfo.InvariantCulture) + "."); }
            private static int End(FormulaNode node) { return node.Start + node.Length; }
            private Token EndToken() { return new Token { Kind = TokenKind.End, Start = _source.Length, Length = 0, Text = String.Empty }; }
        }
    }
}
