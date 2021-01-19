using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.LiveCodingLab.Lexer;
using PowerPointLabs.LiveCodingLab.Lexer.Grammars;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.LiveCodingLab.Service
{
    public class SyntaxHighlightingUploadService
    {
        /// <summary>
        /// Retrieves code from a file and converts to a string for further processing
        /// </summary>
        /// <param name="filePath">filePath of the file containing the code</param>
        /// <returns>parsed code from file in string format</returns>
        public static Tuple<string, List<LexicalRule>, Dictionary<TokenType, Color>> GetColourSchemeFromJson(string filePath)
        {
            string directory = Directory.GetParent(Directory.GetParent(Directory.GetParent(System.AppDomain.CurrentDomain.BaseDirectory).FullName).FullName).FullName + "\\LiveCodingLab\\Service\\";
            // Check if file exists, inform user if file does not exist
            if (!File.Exists(directory + filePath))
            {
                MessageBox.Show(LiveCodingLabText.ErrorInvalidFileName,
                    LiveCodingLabText.ErrorHighlightDifferenceDialogTitle);
                return null;
            }

            string text = File.ReadAllText(directory + filePath);

            JObject customGrammar = JObject.Parse(text);

            string language = (string)customGrammar["language"];

            var grammarRules = customGrammar["rules"].Children();

            List<LexicalRule> rules = new List<LexicalRule>();
            foreach (var rule in grammarRules)
            {
                TokenType type = GetTokenTypeFromString((string)rule["type"]);
                if (type == TokenType.Comment || type == TokenType.Identifier || type == TokenType.String ||
                    type == TokenType.WhiteSpace || type == TokenType.Unknown)
                {
                    string ruleToAdd = (string)rule["regex"];
                    if (ruleToAdd[0] == '@')
                    {
                        ruleToAdd = ruleToAdd.Substring(1);
                        rules.Add(new LexicalRule { Type = type, RegExpression = new Regex(@ruleToAdd) });
                    }
                    else
                    {
                        rules.Add(new LexicalRule { Type = type, RegExpression = new Regex(ruleToAdd) });
                    }
                }
                else if (type == TokenType.Control || type == TokenType.Metadata || type == TokenType.Builtins ||
                         type == TokenType.Keyword)
                {
                    rules.Add(new LexicalRule { Type = type, RegExpression = LexicalRule.WordRegex(rule["regex"].ToObject<string[]>()) });
                }
                else
                {
                    continue;
                }
            }

            var grammarColors = customGrammar["color"].Children();

            Dictionary<TokenType, Color> colorDict = new Dictionary<TokenType, Color>();
            foreach (var tokenType in grammarColors)
            {
                TokenType name = GetTokenTypeFromString((string)tokenType["name"]);
                int alpha = (int)tokenType["alpha"];
                int blue = (int)tokenType["blue"];
                int green = (int)tokenType["green"];
                int red = (int)tokenType["red"];
                colorDict[name] = Color.FromArgb(alpha, blue, green, red);
            }

            return Tuple.Create(language, rules, colorDict);
        }

        public static TokenType GetTokenTypeFromString(string type)
        {
            switch (type)
            {
                case "Comment":
                    return TokenType.Comment;
                case "Number":
                    return TokenType.Number;
                case "String":
                    return TokenType.String;
                case "Operator":
                    return TokenType.Operator;
                case "Delimiter":
                    return TokenType.Delimiter;
                case "Keyword":
                    return TokenType.Keyword;
                case "Builtins":
                    return TokenType.Builtins;
                case "Control":
                    return TokenType.Control;
                case "Identifier":
                    return TokenType.Identifier;
                case "Metadata":
                    return TokenType.Metadata;
                case "WhiteSpace":
                    return TokenType.WhiteSpace;
                case "Unknown":
                    return TokenType.Unknown;
                default:
                    return TokenType.Unknown;
            }
        }
    }
}
