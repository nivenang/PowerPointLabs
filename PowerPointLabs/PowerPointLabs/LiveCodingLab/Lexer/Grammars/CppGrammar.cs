using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;

namespace Lexer.Grammars
{
    public class CppGrammar : IGrammar
    {
        public CppGrammar()
        {
            Rules = new List<LexicalRule>()
            {
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex("^(//.*)") }, // Comment
                new LexicalRule { Type = TokenType.Comment, RegExpression = LexicalRule.WordRegex("#include") },
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex(@"/\*(?:(?!\*/).)*\*/") },
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex(@"/\*(?:(?!\*/).)*") },
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex(@"(?:(?!\*/).)*\*/") },
                new LexicalRule { Type = TokenType.WhiteSpace, RegExpression = new Regex("^\\s") }, // Whitespace

                new LexicalRule 
                { 
                    Type = TokenType.Keyword, 
                    RegExpression = LexicalRule.WordRegex(
                        "asm", "auto", "bool", "break", "case", "catch", "char", "class", "const", "const_cast", "continue",
                        "default", "delete", "do", "double", "dynamic_cast", "else", "enum", "explicit", "export", "extern", "false",
                        "float", "for", "friend", "goto", "if", "inline", "int", "long", "mutable", "namespace", "new",
                        "operator", "private", "protected", "public", "register", "reinterpret_cast", "return", "short", "signed", 
                        "sizeof", "static", "static_cast", "struct", "switch", "template", "this", "throw", "true", "try", "typedef", 
                        "typeid", "typename", "union", "unsigned", "using", "virtual", "void", "volatile", "wchar_t", "while") 
                }, // Keywords
                
                new LexicalRule { Type = TokenType.Identifier, RegExpression = new Regex("^[_A-Za-z][_A-Za-z0-9]*") }, // Identifier
                new LexicalRule { Type = TokenType.String, RegExpression = new Regex("^((@'(?:[^']|'')*'|'(?:\\.|[^\\']|)*('|\\b))|(@\"(?:[^\"]|\"\")*\"|\"(?:\\.|[^\\\"])*(\"|\\b)))", RegexOptions.IgnoreCase) }, // String Marker
                new LexicalRule { Type = TokenType.String, RegExpression = new Regex("(<.*>)") },
                new LexicalRule { Type = TokenType.Unknown, RegExpression = new Regex("^.") }, // Any
            };

            ColorDict = new Dictionary<TokenType, Color>()
            {
                { TokenType.Comment, Color.FromArgb(255, 40, 200, 0) },
                { TokenType.String, Color.FromArgb(255, 0, 0, 171) },
                { TokenType.Builtins, Color.FromArgb(255, 144, 0, 144) },
                { TokenType.Keyword, Color.FromArgb(255, 255, 119, 0) },
                { TokenType.Identifier, Color.FromArgb(0, 0, 0, 0) },
                { TokenType.Delimiter, Color.FromArgb(0, 0, 0, 0) },
                { TokenType.WhiteSpace, Color.FromArgb(0, 0, 0, 0) },
                { TokenType.Number, Color.FromArgb(0, 0, 0, 0) },
                { TokenType.Operator, Color.FromArgb(255, 0, 0, 0) },
                { TokenType.Unknown, Color.FromArgb(0, 0, 0, 0) },
            };
        }

        public string Name
        {
            get { return "Java"; }
        }

        public List<LexicalRule> Rules { get; private set; }

        public Dictionary<TokenType, Color> ColorDict { get; }
    }
}