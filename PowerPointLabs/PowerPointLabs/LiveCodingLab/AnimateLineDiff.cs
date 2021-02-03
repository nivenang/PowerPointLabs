using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.LiveCodingLab.Model;
using PowerPointLabs.LiveCodingLab.Service;
using PowerPointLabs.LiveCodingLab.Utility;
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
        internal const int AnimateLineDiff_MinNoOfShapesRequired = 1;
        internal const string AnimateLineDiff_FeatureName = "Animate Line Diff";
        internal const string AnimateLineDiff_ShapeSupport = "code box";
        internal static readonly string[] AnimateLineDiff_ErrorParameters =
        {
            AnimateLineDiff_FeatureName,
            AnimateLineDiff_MinNoOfShapesRequired.ToString(),
            AnimateLineDiff_ShapeSupport
        };
        public void AnimateLineDiff(List<CodeBoxPaneItem> codeListBox)
        {
            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                
                // Check that there is a slide selected by the user
                if (currentSlide == null)
                {
                    currentSlide = currentPresentation.Slides[currentPresentation.SlideCount - 1];
                }

                // Check that there exists a "before" and "after" code
                if (codeListBox.Count != 2)
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateLineDiffDialogTitle);
                    return;
                }

                CodeBoxPaneItem diffCodeBoxBefore = codeListBox[0];
                CodeBoxPaneItem diffCodeBoxAfter = codeListBox[1];

                FileDiff diffFile;

                // Case 1: Animating differences across a Diff File
                if (diffCodeBoxBefore.CodeBox.IsDiff && diffCodeBoxAfter.CodeBox.IsDiff)
                {
                    if (diffCodeBoxBefore.CodeBox.Text != diffCodeBoxAfter.CodeBox.Text)
                    {
                        MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffWrongCodeSnippet,
                                        LiveCodingLabText.ErrorAnimateLineDiffDialogTitle);
                        return;
                    }

                    List<FileDiff> diffList = CodeBoxFileService.ParseDiff(diffCodeBoxBefore.CodeBox.Text);
                    
                    if (diffList.Count < 1)
                    {
                        MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffMissingCodeSnippet,
                                        LiveCodingLabText.ErrorAnimateLineDiffDialogTitle);
                        return;
                    }

                    diffFile = diffList[0];
                }
                // Case 2: Animating differences across two user-input code snippets by building a diff file
                else if (!diffCodeBoxBefore.CodeBox.IsDiff && !diffCodeBoxAfter.CodeBox.IsDiff)
                {
                    // Check that there exists a "before" code and an "after" code to be animated
                    if (diffCodeBoxBefore.CodeBox.Shape == null || diffCodeBoxAfter.CodeBox.Shape == null)
                    {
                        MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffMissingCodeSnippet,
                                        LiveCodingLabText.ErrorAnimateLineDiffDialogTitle);
                        return;
                    }

                    if (diffCodeBoxBefore.CodeBox.Shape.HasTextFrame == Office.MsoTriState.msoFalse ||
                        diffCodeBoxAfter.CodeBox.Shape.HasTextFrame == Office.MsoTriState.msoFalse)
                    {
                        MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffMissingCodeSnippet,
                                        LiveCodingLabText.ErrorAnimateLineDiffDialogTitle);
                        return;
                    }

                    diffCodeBoxAfter.CodeBox.Shape.Left = diffCodeBoxBefore.CodeBox.Shape.Left;
                    diffCodeBoxAfter.CodeBox.Shape.Top = diffCodeBoxBefore.CodeBox.Shape.Top;
                    diffCodeBoxAfter.CodeBox.Shape.Width = diffCodeBoxBefore.CodeBox.Shape.Width;
                    diffCodeBoxAfter.CodeBox.Shape.Height = diffCodeBoxBefore.CodeBox.Shape.Height;

                    var diff = InlineDiffBuilder.Diff(diffCodeBoxBefore.CodeBox.Text, diffCodeBoxAfter.CodeBox.Text);
                    diffFile = BuildDiffFromText(diffCodeBoxBefore.CodeBox.Text, diffCodeBoxAfter.CodeBox.Text);
                }
                // Default: Inform user that code snippets to be animated do not match up
                else
                {
                    MessageBox.Show(LiveCodingLabText.ErrorAnimateDiffMissingCodeSnippet,
                                    LiveCodingLabText.ErrorAnimateLineDiffDialogTitle);
                    return;
                }

                // Run the Animate Diff algorithm on the diff object
                AnimateDiff(codeListBox, diffFile, false);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimateLineDiff");
                throw;
            }
        }

        /// <summary>
        /// Creates a diff file from two user input code boxes for use in animating line diff
        /// </summary>
        /// <param name="text1">text containing the "before" code</param>
        /// <param name="text2">text containing the "after" code</param>
        /// <returns>diff file containing differences between the two code snippets</returns>
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
