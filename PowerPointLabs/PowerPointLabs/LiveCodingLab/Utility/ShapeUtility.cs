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
        /// Insert default callout box shape to slide. 
        /// Precondition: shape with shapeName must not exist in slide before applying the method
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="codeText">Content in Callout Shape</param>
        /// <returns>generated callout shape</returns>
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
            Shape codeShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, 10, 500, 250);
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
        /// Replace original text in `shape` with `text`
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        public static CodeBox ReplaceTextForShape(CodeBox codeBox)
        {
            if (!codeBox.Slide.HasShapeWithSameName(string.Format(LiveCodingLabText.CodeBoxShapeNameFormat, codeBox.Id)))
            {
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
            else
            {
                shapeInSlide.TextFrame.TextRange.Text = codeBox.Text;
            }
            codeBox.Shape = shapeInSlide;
            return codeBox;
        }
    }
}
