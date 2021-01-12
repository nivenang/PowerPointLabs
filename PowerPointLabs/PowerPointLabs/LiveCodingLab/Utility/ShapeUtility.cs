﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using Highlight;
using Highlight.Engines;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.LiveCodingLab.Lexer;
using PowerPointLabs.LiveCodingLab.Lexer.Grammars;
using PowerPointLabs.LiveCodingLab.Model;
using PowerPointLabs.LiveCodingLab.Service;
using PowerPointLabs.LiveCodingLab.Views;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.LiveCodingLab.Utility
{
    public class ShapeUtility
    {
#pragma warning disable 0618
        /// <summary>
        /// Insert code box text to specified slide. 
        /// Precondition: shape with codeBox.shapeName must not exist in slide before applying the method
        /// </summary>
        /// <param name="slide">Slide to insert textbox into</param>
        /// <param name="codeBox">CodeBox object containing the code snippet</param>
        /// <returns>generated code text box</returns>
        public static CodeBox InsertCodeBoxToSlide(PowerPointSlide slide, CodeBox codeBox)
        {
            string textToInsert;

            if (codeBox.IsFile)
            {
                textToInsert = CodeBoxFileService.GetCodeFromFile(codeBox.Text);
            }
            else
            {
                textToInsert = codeBox.Text;
            }

            // Create textbox in the specified slide
            Shape codeShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 170, 100, 700, 250);
            
            // Insert text into the created textbox
            if (textToInsert != null && textToInsert != "")
            {
                codeShape.TextFrame.TextRange.Text = textToInsert;
            }
            codeShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            codeShape.TextFrame.WordWrap = MsoTriState.msoTrue;
            codeShape.TextFrame.TextRange.Font.Size = LiveCodingLabSettings.codeFontSize;
            codeShape.TextFrame.TextRange.Font.Name = LiveCodingLabSettings.codeFontType;
            codeShape.TextFrame.TextRange.Font.Color.RGB = LiveCodingLabSettings.codeTextColor.ToArgb();
            codeShape.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentLeft;
            codeShape.Name = string.Format(LiveCodingLabText.CodeBoxShapeNameFormat, codeBox.Id);
            codeBox.Slide = slide;

            // Highlight the syntax of the code in the textbox
            if (codeBox.IsFile)
            {
                codeBox.Shape = HighlightSyntax(codeShape, slide, codeBox.FileText);
            }
            else
            {
                codeBox.Shape = HighlightSyntax(codeShape, slide);
            }
            codeBox.ShapeName = codeShape.Name;
            return codeBox;
        }

        /// <summary>
        /// Insert code box text of a diff file to slide. 
        /// Precondition: shape with codeBox.shapeName must not exist in slide before applying the method
        /// </summary>
        /// <param name="slide">Slide to insert textbox into</param>
        /// <param name="codeBox">CodeBox object containing the code snippet</param>
        /// <param name="diff">diff file containing differences in code across two code snippets</param>
        /// <returns>generated code text box</returns>
        public static CodeBox InsertDiffCodeBoxToSlide(PowerPointSlide slide, CodeBox codeBox, FileDiff diff)
        {
            string textToInsert = CodeBoxFileService.ConvertFileDiffToString(diff)[codeBox.DiffIndex];

            // Create textbox in the specified slide
            Shape codeShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 170, 100, 700, 250);

            // Insert text into the created textbox
            if (textToInsert != null && textToInsert != "")
            {
                codeShape.TextFrame.TextRange.Text = textToInsert;
            }
            codeShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            codeShape.TextFrame.WordWrap = MsoTriState.msoTrue;
            codeShape.TextFrame.TextRange.Font.Size = LiveCodingLabSettings.codeFontSize;
            codeShape.TextFrame.TextRange.Font.Name = LiveCodingLabSettings.codeFontType;
            codeShape.TextFrame.TextRange.Font.Color.RGB = LiveCodingLabSettings.codeTextColor.ToArgb();
            codeShape.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentLeft;
            codeShape.Name = string.Format(LiveCodingLabText.CodeBoxShapeNameFormat, codeBox.Id);
            codeBox.Slide = slide;

            // Highlight the syntax of the code in the textbox
            codeBox.Shape = HighlightSyntax(codeShape, slide);
            codeBox.ShapeName = codeShape.Name;
            return codeBox;
        }

        /// <summary>
        /// Insert a textbox containing details of all code boxes into specified slide for storage purposes. 
        /// </summary>
        /// <param name="slide">Slide to insert textbox into</param>
        /// <param name="shapeName">Name of text box used for the storage</param>
        /// <param name="text">Encoded text for storage of code boxes</param>
        /// <returns>generated text box containing encoded storage text</returns>
        public static Shape InsertStorageCodeBoxToSlide(PowerPointSlide slide, string shapeName, string text)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;

            // Create textbox in the specified slide
            Shape storageBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0,
                slideWidth, 100);
            storageBox.Name = shapeName;
            storageBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            storageBox.TextFrame.TextRange.Text = text;
            storageBox.TextFrame.WordWrap = MsoTriState.msoTrue;
            storageBox.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            storageBox.TextFrame.TextRange.Font.Size = 12;
            storageBox.Fill.BackColor.RGB = 0xffffff;
            storageBox.Fill.Transparency = 0.2f;
            storageBox.TextFrame.TextRange.Font.Color.RGB = 0;
            storageBox.Visible = MsoTriState.msoFalse;

            return storageBox;
        }

        /// <summary>
        /// Replace original text in CodeBox shape on slide with the updated CodeBox text
        /// </summary>
        /// <param name="codeBox">Code box to be updated</param>
        /// <returns>updated codeBox containing the new text</returns>
        public static CodeBox ReplaceTextForShape(CodeBox codeBox)
        {
            // Insert a new code box into the slide if code box does not exist in the slide
            try
            {
                if (codeBox.Shape == null && !codeBox.Slide.HasShapeWithSameName(string.Format(LiveCodingLabText.CodeBoxShapeNameFormat, codeBox.Id)))
                {
                    if (codeBox.IsDiff && codeBox.DiffIndex >= 0)
                    {
                        return InsertDiffCodeBoxToSlide(codeBox.Slide, codeBox, CodeBoxFileService.ParseDiff(codeBox.Text)[0]);
                    }
                    return InsertCodeBoxToSlide(codeBox.Slide, codeBox);
                }
            } 
            catch (COMException)
            {
                if (codeBox.IsDiff && codeBox.DiffIndex >= 0)
                {
                    return InsertDiffCodeBoxToSlide(PowerPointCurrentPresentationInfo.CurrentSlide, codeBox, CodeBoxFileService.ParseDiff(codeBox.Text)[0]);
                }
                return InsertCodeBoxToSlide(PowerPointCurrentPresentationInfo.CurrentSlide, codeBox);
            }

            // Update the code in the slide to the new code
            Shape shapeInSlide = codeBox.Shape;
            shapeInSlide.TextFrame.TextRange.Font.Name = LiveCodingLabSettings.codeFontType;
            shapeInSlide.TextFrame.TextRange.Font.Size = LiveCodingLabSettings.codeFontSize;
            shapeInSlide.TextFrame.TextRange.Font.Color.RGB = LiveCodingLabSettings.codeTextColor.ToArgb();
            
            // Highlights the syntax of the updated code
            if (codeBox.IsFile)
            {
                shapeInSlide.TextFrame.TextRange.Text = CodeBoxFileService.GetCodeFromFile(codeBox.Text);
                codeBox.Shape = HighlightSyntax(shapeInSlide, codeBox.Slide, codeBox.FileText);
            }
            else if (codeBox.IsDiff)
            {
                shapeInSlide.TextFrame.TextRange.Text = CodeBoxFileService.ConvertFileDiffToString(CodeBoxFileService.ParseDiff(codeBox.Text)[0])[codeBox.DiffIndex];
                codeBox.Shape = HighlightSyntax(shapeInSlide, codeBox.Slide);
            }
            else
            {
                shapeInSlide.TextFrame.TextRange.Text = codeBox.Text;
                codeBox.Shape = HighlightSyntax(shapeInSlide, codeBox.Slide);
            }
            codeBox.ShapeName = shapeInSlide.Name;
            return codeBox;
        }

        /// <summary>
        /// Highlights the syntax of the code box according to user's preferences 
        /// </summary>
        /// <param name="shape">Shape containing the code in the slide</param>
        /// <param name="slide">Slide that contains the shape to be highlighted</param>
        /// <param name="filePath">filePath of the code if it exists</param>
        /// <returns>generated code box with highlighted syntax</returns>
        private static Shape HighlightSyntax(Shape shape, PowerPointSlide slide, string filePath="")
        {
            // Break down entire code into individual lines
            Shape shapeToProcess = ConvertTextToParagraphs(shape);
           
            // Generates the relevant grammar to highlight the syntax based on user specifications
            Dictionary<string, Type> stringToGrammar = new Dictionary<string, Type>
            {
                { "Java", typeof(JavaGrammar) },
                { "Python", typeof(PythonGrammar) },
                { "C", typeof(CGrammar) },
                { "C++", typeof(CppGrammar) },
            };

            Dictionary<string, Type> fileToGrammar = new Dictionary<string, Type>
            {
                { "java", typeof(JavaGrammar) },
                { "py", typeof(PythonGrammar) },
                { "c", typeof(CGrammar) },
                { "cpp", typeof(CppGrammar) },
                { "cxx", typeof(CppGrammar) },
                { "cc", typeof(CppGrammar) },
                { "hpp", typeof(CppGrammar) },
                { "hxx", typeof(CppGrammar) },
            };

            IGrammar grammar;
            Tokenizer lexer;

            
            try
            {
                // Automatically select the best grammar for syntax highlighting if code box is a file
                if (filePath != "" && filePath.LastIndexOf('.') >= 0 && fileToGrammar.ContainsKey(filePath.Substring(filePath.LastIndexOf('.')+1)))
                {
                    grammar = (IGrammar)Activator.CreateInstance(fileToGrammar[filePath.Substring(filePath.LastIndexOf('.') + 1)]);
                }
                // Use default grammar if no grammar is specified by user
                else if (LiveCodingLabSettings.language.Equals("None"))
                {
                    string keyWords = "(abstract|as|base|bool|break|byte|case|catch|char|checked|class|const|continue|decimal|default|delegate|do|double|else|enum|event|explicit|extern|false|finally|fixed|float|for|" +
                        "foreach|goto|if|implicit|import|int|in|interface|internal|is|lock|long|namespace|new|null|object|operator|out|override|params|private|protected|public|readonly|ref|return|sbyte|sealed|short|sizeof|stackalloc|static|" +
                        "string|struct|switch|this|throw|true|try|typeof|uint|ulong|unchecked|unsafe|ushort|using|virtual|volatile|void|while)";

                    TextRange textRange = shapeToProcess.TextFrame.TextRange;

                    foreach (TextRange paragraph in textRange.Paragraphs())
                    {
                        foreach (Match match in Regex.Matches(paragraph.Text, @"\b" + keyWords + @"\b"))
                        {
                            paragraph.Characters(match.Index + 1, match.Length).Font.Color.RGB = Color.Red.ToArgb();
                        }
                    }
                    return shapeToProcess;
                }
                // Use user specified grammar if no filepath exists
                else
                {
                    grammar = (IGrammar)Activator.CreateInstance(stringToGrammar[LiveCodingLabSettings.language]);
                }
                lexer = new Tokenizer(grammar);
            }
            catch (NullReferenceException)
            {
                return shapeToProcess;
            }
            catch (ArgumentNullException)
            {
                return shapeToProcess;
            }

            // Colourise each token based on its token type
            foreach (TextRange paragraph in shapeToProcess.TextFrame.TextRange.Paragraphs())
            {
                foreach (var token in lexer.Tokenize(paragraph.Text))
                {
                    Color color = grammar.ColorDict[token.Type];
                    paragraph.Characters(token.StartIndex + 1, token.Length).Font.Color.RGB = color.ToArgb();
                }
            }

            return shapeToProcess;
        }

        /// <summary>
        /// Converts a chunk of code into individual lines for easier parsing
        /// </summary>
        /// <param name="shape">Shape containing the code in the slide</param>
        /// <returns>shape containing codes broken down into lines</returns>
        private static Shape ConvertTextToParagraphs(Shape shape)
        {
            TextRange codeText = shape.TextFrame.TextRange;
            string textWithParagraphs = "";

            // Create a new line in the code box for every newline encountered
            foreach (TextRange line in codeText.Lines())
            {
                if (line.Text.Contains("\r\n") || line.Text == "")
                {
                    continue;
                }
                else if (line.Text.Contains("\r") && !line.Text.Contains("\n"))
                {
                    textWithParagraphs += line.Text + "\n";
                }
                else if (line.Text.Contains("\n") && !line.Text.Contains("\r"))
                {
                    textWithParagraphs += line.Text.Replace("\n", "\r\n");
                }
                else
                {
                    textWithParagraphs += line.Text + "\r\n";
                }
            }
            shape.TextFrame.TextRange.Text = textWithParagraphs;
            return shape;
        }
    }
}
