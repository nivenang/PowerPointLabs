using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AnimationLab;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.LiveCodingLab.Service;
using PowerPointLabs.LiveCodingLab.Views;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.LiveCodingLab
{
    public partial class LiveCodingLabMain
    {
#pragma warning disable 0618
        internal const int AnimateNewLines_MinNoOfShapesRequired = 1;
        internal const string AnimateNewLines_FeatureName = "Animate New Lines";
        internal const string AnimateNewLines_ShapeSupport = "code box";
        internal static readonly string[] AnimateNewLines_ErrorParameters =
        {
            AnimateNewLines_FeatureName,
            AnimateNewLines_MinNoOfShapesRequired.ToString(),
            AnimateNewLines_ShapeSupport
        };

        private static float fontScale = 4.5f;

        public void AnimateNewLines(List<CodeBoxPaneItem> listCodeBox)
        {
            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

                if (currentSlide == null || currentSlide.Index == PowerPointPresentation.Current.SlideCount)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesWrongSlide,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                PowerPointSlide nextSlide = PowerPointPresentation.Current.Slides[currentSlide.Index];

                //Get shapes to consider for animation
                CodeBoxPaneItem currentSlideCodeBox = listCodeBox[0];
                CodeBoxPaneItem nextSlideCodeBox = listCodeBox[1];

                PowerPoint.Shape currentSlideShape = currentSlideCodeBox.CodeBox.Shape;

                // Check that there exists a "before" code and an "after" code to be animated
                if (currentSlideCodeBox.CodeBox.Shape == null || nextSlideCodeBox.CodeBox.Shape == null)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                if (currentSlideCodeBox.CodeBox.Shape.HasTextFrame == Office.MsoTriState.msoFalse ||
                    nextSlideCodeBox.CodeBox.Shape.HasTextFrame == Office.MsoTriState.msoFalse)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateNewLinesMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateNewLinesDialogTitle);
                    return;
                }

                nextSlideCodeBox.CodeBox.Shape.Left = currentSlideCodeBox.CodeBox.Shape.Left;
                nextSlideCodeBox.CodeBox.Shape.Top = currentSlideCodeBox.CodeBox.Shape.Top;
                nextSlideCodeBox.CodeBox.Shape.Width = currentSlideCodeBox.CodeBox.Shape.Width;
                nextSlideCodeBox.CodeBox.Shape.Height = currentSlideCodeBox.CodeBox.Shape.Height;

                var diff = InlineDiffBuilder.Diff(currentSlideCodeBox.CodeBox.Text, nextSlideCodeBox.CodeBox.Text);
                FileDiff diffFile = BuildDiffFromText(currentSlideCodeBox.CodeBox.Text, nextSlideCodeBox.CodeBox.Text);
                AnimateDiffByLine(listCodeBox, diffFile);

                currentSlideCodeBox.CodeBox.Slide = currentSlide;
                currentSlideCodeBox.CodeBox.Shape = currentSlideShape;
                nextSlideCodeBox.CodeBox.Slide = nextSlide;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateNewLines");
                throw;
            }
        }

        private static FileDiff BuildDiffFromText(string text1, string text2)
        {
            var diff = InlineDiffBuilder.Diff(text1, text2);
            string diffFile = "--- /path/to/file1	2020-09-28 23:30:39.942229878 -0800\r\n" +
                "+++ /path/to/file2  2020-09-28 23:30:50.442260588 -0800\r\n" +
                "@@ -1,1 +1,1 @@\r\n";

            foreach (var line in diff.Lines)
            {
                switch (line.Type)
                {
                    case ChangeType.Inserted:
                        diffFile += AppendLineEnd("+" + line.Text);
                        break;
                    case ChangeType.Deleted:
                        diffFile += AppendLineEnd("-" + line.Text);
                        break;
                    default:
                        diffFile += AppendLineEnd(" " + line.Text);
                        break;
                }
            }

            List<FileDiff> diffList = Diff.Parse(diffFile, Environment.NewLine).ToList();
            return diffList[0];
        }
    }
}
