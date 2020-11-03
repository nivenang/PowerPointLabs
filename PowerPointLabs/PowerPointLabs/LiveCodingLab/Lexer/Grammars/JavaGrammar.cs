// Copyright (c) Adnan Umer. All rights reserved. Follow me @aztnan
// Email: aztnan@outlook.com
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;

namespace Lexer.Grammars
{
    public class JavaGrammar : IGrammar
    {
        public JavaGrammar()
        {
            Rules = new List<LexicalRule>()
            {
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex("^(//.*)") }, // Comment
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex(@"/\*(?:(?!\*/).)*\*/") },
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex(@"/\*(?:(?!\*/).)*") },
                new LexicalRule { Type = TokenType.Comment, RegExpression = new Regex(@"(?:(?!\*/).)*\*/") },
                new LexicalRule { Type = TokenType.WhiteSpace, RegExpression = new Regex("^\\s") }, // Whitespace
                new LexicalRule { Type = TokenType.Operator, RegExpression = new Regex("^[\\+\\-\\*/%&|\\^~<>!]") }, // Single Char Operator
                new LexicalRule { Type = TokenType.Operator, RegExpression = new Regex("^((==)|(!=)|(<=)|(>=)|(<>)|(<<)|(>>)|(//)|(\\*\\*))") }, // Double Char Operator
                new LexicalRule { Type = TokenType.Delimiter, RegExpression = new Regex("^[\\(\\)\\[\\]\\{\\}@,:`=;\\.]") }, // Single Delimiter
                new LexicalRule { Type = TokenType.Delimiter, RegExpression = new Regex("^((\\+=)|(\\-=)|(\\*=)|(%=)|(/=)|(&=)|(\\|=)|(\\^=))") }, // Double Char Operator
                new LexicalRule { Type = TokenType.Delimiter, RegExpression = new Regex("^((//=)|(>>=)|(<<=)|(\\*\\*=))") }, // Triple Delimiter

                new LexicalRule 
                { 
                    Type = TokenType.Keyword, 
                    RegExpression = LexicalRule.WordRegex(
                        "abstract", "assert", "boolean", "break", "byte", "case", "catch", "char", "class", "const",
                        "continue", "default", "do", "double", "else", "enum", "extends", "final", "finally", "float",
                        "for", "goto", "if", "implements", "import", "instanceof", "int", "interface", "long", "native",
                        "new", "package", "private", "protected", "public", "return", "short", "static", "strictfp", "super", 
                        "switch", "synchronized", "this", "throw", "throws", "transient", "try", "void", "volatile", "while") 
                }, // Keywords
                
                new LexicalRule { Type = TokenType.Identifier, RegExpression = new Regex("^[_A-Za-z][_A-Za-z0-9]*") }, // Identifier
                new LexicalRule { Type = TokenType.String, RegExpression = new Regex("^((@'(?:[^']|'')*'|'(?:\\.|[^\\']|)*('|\\b))|(@\"(?:[^\"]|\"\")*\"|\"(?:\\.|[^\\\"])*(\"|\\b)))", RegexOptions.IgnoreCase) }, // String Marker
                
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