using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;

namespace PowerPointLabs.LiveCodingLab.Lexer.Grammars
{
    public class CGrammar : IGrammar
    {
        public CGrammar()
        {
            Rules = new List<LexicalRule>()
            {
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex("^(//.*)") }, // Comment
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex(@"/\*(?:(?!\*/).)*\*/") },
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex(@"/\*(?:(?!\*/).)*") },
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex(@"(?:(?!\*/).)*\*/") },
                //new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex("^(\\*.*)") },
                new LexicalRule { Type = TokenType.Metadata, RegExpression = LexicalRule.WordRegex("#include", "#define") },
                new LexicalRule { Type = TokenType.WhiteSpace, RegExpression = new Regex("^\\s") }, // Whitespace
                new LexicalRule
                {
                    Type = TokenType.Control,
                    RegExpression = LexicalRule.WordRegex("break", "case", "continue", "default", "do", "else", "for", "goto",
                        "if", "return", "switch", "while")
                },
                new LexicalRule 
                { 
                    Type = TokenType.Keyword, 
                    RegExpression = LexicalRule.WordRegex(
                        "auto", "char", "const", "double", "enum", "extern",
                        "float", "int", "long", "NULL", "register", "short", "signed", "size_t", "sizeof", 
                        "static", "struct", "typedef", "union", "unsigned", "void", "volatile", "_Packed") 
                }, // Keywords
                new LexicalRule
                {
                    Type = TokenType.Builtins,
                    RegExpression = LexicalRule.WordRegex(
                    "calloc", "free", "malloc", "realloc", "va_list", "va_start", "va_arg", "va_end", "FILE",
                    "getchar", "putchar", "gets", "puts", "scanf", "printf", "fscanf", "fputc", "fputs",
                    "fclose", "fgets", "fgetc", "fread", "fwrite", "fprintf")
                },
                new LexicalRule { Type = TokenType.Identifier, RegExpression = new Regex("^[_A-Za-z][_A-Za-z0-9]*") }, // Identifier
                new LexicalRule { Type = TokenType.String, RegExpression = new Regex("^((@'(?:[^']|'')*'|'(?:\\.|[^\\']|)*('|\\b))|(@\"(?:[^\"]|\"\")*\"|\"(?:\\.|[^\\\"])*(\"|\\b)))", RegexOptions.IgnoreCase) }, // String Marker
                new LexicalRule { Type = TokenType.String, RegExpression = new Regex("(<.*>)") },
                new LexicalRule { Type = TokenType.Unknown, RegExpression = new Regex("^.") }, // Any
            };

            ColorDict = new Dictionary<TokenType, Color>()
            {
                { TokenType.Comment, Color.FromArgb(128, 0, 128, 0) },
                { TokenType.String, Color.FromArgb(128, 21, 21, 163) },
                { TokenType.Builtins, Color.FromArgb(128, 0, 0, 0) },
                { TokenType.Keyword, Color.FromArgb(128, 255, 0, 0) },
                { TokenType.Identifier, Color.FromArgb(0, 0, 0, 0) },
                { TokenType.Delimiter, Color.FromArgb(0, 0, 0, 0) },
                { TokenType.WhiteSpace, Color.FromArgb(0, 0, 0, 0) },
                { TokenType.Number, Color.FromArgb(0, 0, 0, 0) },
                { TokenType.Operator, Color.FromArgb(255, 0, 0, 0) },
                { TokenType.Metadata, Color.FromArgb(128, 128, 128, 128) },
                { TokenType.Control, Color.FromArgb(128, 196, 0, 143) },
                { TokenType.Unknown, Color.FromArgb(0, 0, 0, 0) },
            };
        }

        public string Name
        {
            get { return "C"; }
        }

        public List<LexicalRule> Rules { get; private set; }

        public Dictionary<TokenType, Color> ColorDict { get; }
    }
}