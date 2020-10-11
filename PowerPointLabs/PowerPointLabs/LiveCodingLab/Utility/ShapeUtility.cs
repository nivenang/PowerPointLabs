using System;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
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
        /// Insert code box text to slide. 
        /// Precondition: shape with codeBox.shapeName must not exist in slide before applying the method
        /// </summary>
        /// <param name="slide"></param>
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
            Shape codeShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 15, 15, 700, 250);
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
            codeBox.Shape = codeShape;
            codeBox.ShapeName = codeShape.Name;
            return codeBox;
        }

        /// <summary>
        /// Insert code box text to slide. 
        /// Precondition: shape with codeBox.shapeName must not exist in slide before applying the method
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="codeBox">CodeBox object containing the code snippet</param>
        /// <returns>generated code text box</returns>
        public static CodeBox InsertDiffCodeBoxToSlide(PowerPointSlide slide, CodeBox codeBox, FileDiff diff)
        {
            string textToInsert = CodeBoxFileService.ConvertFileDiffToString(diff)[codeBox.DiffIndex];
            Shape codeShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 15, 15, 700, 250);
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
            codeBox.Shape = codeShape;
            codeBox.ShapeName = codeShape.Name;
            return codeBox;
        }

        public static Shape InsertStorageCodeBoxToSlide(PowerPointSlide slide, string shapeName, string text)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;

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
        /// <param name="codeBox"></param>
        /// <returns>updated codeBox containing the new text</returns>
        public static CodeBox ReplaceTextForShape(CodeBox codeBox)
        {
            if (!codeBox.Slide.HasShapeWithSameName(string.Format(LiveCodingLabText.CodeBoxShapeNameFormat, codeBox.Id)))
            {
                if (codeBox.IsDiff && codeBox.DiffIndex >= 0)
                {
                    return InsertDiffCodeBoxToSlide(codeBox.Slide, codeBox, CodeBoxFileService.ParseDiff(codeBox.Text)[0]);
                }
                return InsertCodeBoxToSlide(codeBox.Slide, codeBox);
            }
            Shape shapeInSlide = codeBox.Shape;
            shapeInSlide.TextFrame.TextRange.Font.Name = LiveCodingLabSettings.codeFontType;
            shapeInSlide.TextFrame.TextRange.Font.Size = LiveCodingLabSettings.codeFontSize;
            shapeInSlide.TextFrame.TextRange.Font.Color.RGB = LiveCodingLabSettings.codeTextColor.ToArgb();
            if (codeBox.IsFile)
            {
                shapeInSlide.TextFrame.TextRange.Text = CodeBoxFileService.GetCodeFromFile(codeBox.Text);
            }
            else if (codeBox.IsDiff)
            {
                shapeInSlide.TextFrame.TextRange.Text = CodeBoxFileService.ConvertFileDiffToString(CodeBoxFileService.ParseDiff(codeBox.Text)[0])[codeBox.DiffIndex];
            }
            else
            {
                shapeInSlide.TextFrame.TextRange.Text = codeBox.Text;
            }
            codeBox.ShapeName = shapeInSlide.Name;
            codeBox.Shape = shapeInSlide;
            return codeBox;
        }
    }
}
